#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
配置持久化
==========
把用户配置存到磁盘，下次启动自动恢复。

存哪里
------
**不能存 exe 旁边。** PyInstaller `--onefile` 运行时会把自身解压到
`%TEMP%\_MEIxxxxxx`，那个目录随进程结束而消失。存在旁边的配置下次启动
就没了。所以走标准用户配置目录：

    %APPDATA%\\WeChatFileSender\\config.json
    (macOS/Linux: ~/.config/wechat-file-sender/config.json)

存什么、不存什么
----------------
只存**下次还能继续干活**的东西。运行时状态一律不存——重启后它们
要么没有意义，要么是错的：

  存：folder_path / mappings / enable_msg / msg / max_retries / recent_folders
  不存：
    sending        重启后必然是 False
    logs           跑完就没了，存它只会让下次启动看到一堆旧日志
    progress/result同上，同属单次运行的产物
    active_queue   含内存对象引用，**根本无法 JSON 序列化**
    needs_review   需要人去微信里核对，重启后语境已丢失，留着误导

安全阀：路径校验
---------------
配置文件是可被编辑的。万一里面写了个别人的敏感目录，下次启动直接
指向那儿就会出问题。所以加载时校验 `folder_path` 仍然存在且是目录，
否则丢弃该字段并记录原因。

写入是原子的
------------
先写 `.tmp` 再 `os.replace`。直接覆写时若中途崩溃/断电，
会留下半截 JSON，下次启动就整个读不出来。`os.replace` 在同一文件系统
内是原子的rename。
"""

import json
import os
import sys
import tempfile
import threading
import time
from typing import Any, Dict, List, Optional

# 配置文件格式版本。将来结构变了可据此做迁移
CONFIG_VERSION = 1

# ── 允许持久化的字段（白名单）──────────────────────────────
# 新增字段必须显式加进来，避免不小心把运行时状态存进去
PERSIST_FIELDS = (
    'folder_path',      # 待发送文件夹
    'mappings',         # [(关键词, 备注名/昵称), ...]
    'enable_msg',       # 是否附带消息
    'msg',# 附言模板
    'max_retries',      # 失败重试次数
    'recent_folders',   # 最近使用文件夹列表（给下拉框用）
    'window_geometry',  # 窗口位置尺寸
)

# 明确不该持久化的字段（写进文件仅作自文档，便于排查）
NEVER_PERSIST = (
    'sending', 'stop_requested', 'logs', 'progress', 'result',
    'stats', 'queue', 'needs_review', 'active_queue',
    'excel_uploaded', 'matched', 'grouped',
)

# mappings 是 tuple 列表，JSON 只能存 list。读取时转回 tuple 以匹配原行为
_TUPLE_FIELDS = {'mappings'}


def _default_config_dir() -> str:
    """
    按平台返回用户配置目录。

    尊重环境变量 WFS_CONFIG_DIR —— 测试用它指向临时目录，
    避免测试往用户真实的 %APPDATA% 写配置。
    """
    override = os.environ.get('WFS_CONFIG_DIR')
    if override:
        return override

    if sys.platform == 'win32':
        base = os.environ.get('APPDATA') or os.path.expanduser('~')
        return os.path.join(base, 'WeChatFileSender')
    # macOS / Linux
    base = (os.environ.get('XDG_CONFIG_HOME')
            or os.path.join(os.path.expanduser('~'), '.config'))
    return os.path.join(base, 'wechat-file-sender')


def default_config_path() -> str:
    return os.path.join(_default_config_dir(), 'config.json')


class ConfigStore:
    """
    配置的加载/保存。线程安全，写入带防抖。

    典型用法：
        cfg = ConfigStore()
        cfg.mark_dirty()                 # 请求保存（防抖后落盘）
        cfg.save_now()                   # 立即落盘（关键时机）
        data = cfg.load()                # 读取并校验
    """

    def __init__(self, path: Optional[str] = None,
                 debounce_seconds: float = 1.5,
                 enabled: bool = True):
        """
        path   配置文件路径，None则用默认位置
        debounce_seconds  连续修改的最小落盘间隔，避免高频写盘
        enabled False 时所有保存操作变成no-op（测试用）
        """
        self.path = path or default_config_path()
        self.debounce_seconds = debounce_seconds
        self.enabled = enabled

        self._lock = threading.RLock()

        self._last_save = 0.0
        self._timer: Optional[threading.Timer] = None
        self._provider = None      #绑定的 state，flush 时从中取值
        # 额外的字段过滤器: field_name -> (state) -> bool
        # 用于「值虽然合法但此刻不该存」的动态场景
        # 例：mappings 只有手动输入的才值得存，上传的 Excel 路径重启后可能失效
        self._field_filters: Dict[str, Any] = {}
        # 过滤器报错: field_name -> 原因。单独存，不与 last_error 混用
        # （last_error 会被后续的成功写入清空，而这个需要留痕）
        self.filter_errors: Dict[str, str] = {}

        # 诊断信息
        self.last_error: Optional[str] = None
        self.load_warnings: List[str] = []

    def add_field_filter(self, field: str, predicate):
        """
        注册字段过滤器。predicate(state) -> bool，False 表示本次不存该字段。
        """
        self._field_filters[field] = predicate

    # ══════════════════════════════════════════════════════════
    #  加载
    # ══════════════════════════════════════════════════════════

    def load(self) -> Dict[str, Any]:
        """
        读取配置。不存在或损坏时返回 {}，绝不抛异常——
        配置文件读不出来不该阻止程序启动。
        """
        self.load_warnings = []
        if not os.path.isfile(self.path):
            return {}

        try:
            with open(self.path, 'r', encoding='utf-8') as f:
                raw = json.load(f)
        except json.JSONDecodeError as e:
            self.last_error = f'配置文件格式错误: {e}'
            self.load_warnings.append(
                '配置文件格式损坏，已忽略并使用默认设置')
            self._backup_corrupt()
            return {}
        except OSError as e:
            self.last_error = f'配置文件读取失败: {e}'
            self.load_warnings.append('配置文件读取失败，已使用默认设置')
            return {}

        if not isinstance(raw, dict):
            self.load_warnings.append('配置文件内容不是对象，已忽略')
            return {}

        # 版本迁移钩子（当前只有 v1）
        version = raw.get('version', 0)
        if version > CONFIG_VERSION:
            self.load_warnings.append(
                f'配置版本 {version} 高于当前支持的 {CONFIG_VERSION}，可能来自更新版本')
        raw = self._migrate(raw, version)

        settings = raw.get('settings', {})
        if not isinstance(settings, dict):
            self.load_warnings.append('settings 字段格式错误，已忽略')
            return {}

        clean = self._filter_and_validate(settings)
        return clean

    def _migrate(self, raw: Dict[str, Any], version: int) -> Dict[str, Any]:
        """版本迁移。当前只有 v1 → 无需处理，保留接口。"""
        return raw

    def _filter_and_validate(self, settings: Dict[str, Any]) -> Dict[str, Any]:
        """按白名单过滤并做类型/路径校验。"""
        clean: Dict[str, Any] = {}

        for key in PERSIST_FIELDS:
            if key not in settings:
                continue
            value = settings[key]

            # 路径类字段：必须仍然存在
            if key == 'folder_path':
                if not isinstance(value, str) or not value:
                    continue
                if not os.path.isdir(value):
                    self.load_warnings.append(
                        f'已保存的文件夹不存在，已忽略: {value}')
                    continue
                clean[key] = os.path.normpath(value)

            # 映射表：[(关键词, 名字), ...]
            elif key == 'mappings':
                mappings = self._sanitize_mappings(value)
                if mappings:
                    clean[key] = mappings

            # 重试次数：夹到合法区间
            elif key == 'max_retries':
                try:
                    n = int(value)
                    clean[key] = max(1, min(5, n))
                except (TypeError, ValueError):
                    pass

            # 布尔
            elif key in ('enable_msg',):
                clean[key] = bool(value)

            # 字符串
            elif key in ('msg', 'window_geometry'):
                if isinstance(value, str):
                    clean[key] = value

            # 最近文件夹列表：逐个校验存在性，保留最多 10 个
            elif key == 'recent_folders':
                folders = []
                if isinstance(value, list):
                    for p in value[:10]:
                        if isinstance(p, str) and os.path.isdir(p):
                            folders.append(os.path.normpath(p))
                if folders:
                    clean[key] = folders

        return clean

    @staticmethod
    def _sanitize_mappings(value: Any) -> List:
        """把 mappings 规整成 [(kw, name), ...]，丢弃畸形项。"""
        out = []
        if not isinstance(value, list):
            return out
        for item in value:
            # JSON 读回来是 [['kw','name'], ...]
            if isinstance(item, (list, tuple)) and len(item) == 2:
                kw, nm = item[0], item[1]
                if kw is not None and nm is not None:
                    kw, nm = str(kw).strip(), str(nm).strip()
                    if kw and nm:
                        out.append((kw, nm))
            elif isinstance(item, dict):
                # 容错：也接受 {"keyword":..,"name":..} 形式
                kw = str(item.get('keyword', '')).strip()
                nm = str(item.get('name', '')).strip()
                if kw and nm:
                    out.append((kw, nm))
        return out

    def _backup_corrupt(self):
        """把损坏的配置改名备份，方便用户找回，也避免反复报错。"""
        try:
            stamp = time.strftime('%Y%m%d_%H%M%S')
            bak = f'{self.path}.corrupt-{stamp}'
            os.replace(self.path, bak)
            self.load_warnings.append(f'原配置已备份为 {os.path.basename(bak)}')
        except OSError:
            pass

    # ══════════════════════════════════════════════════════════
    #  保存
    # ══════════════════════════════════════════════════════════

    def build_payload(self, state) -> Dict[str, Any]:
        """从 state 提取可持久化字段。state 可以是 dict 或 StateStore。"""
        settings: Dict[str, Any] = {}
        for key in PERSIST_FIELDS:
            try:
                value = state[key] if not hasattr(state, 'get') else state.get(key)
            except (KeyError, TypeError):
                continue
            # 注意：不能用 `if not value` —— False/0 是有意义的配置值
            #（如 enable_msg=False），只有 None 和空容器才算「无内容」
            if value is None:
                continue
            # 空字符串 / 空列表视为未设置，不落盘（避免存一堆空值）
            if isinstance(value, (str, list, tuple, dict)) and len(value) == 0:
                continue
            # 动态过滤：这个字段此刻该不该存
            filt = self._field_filters.get(key)
            if filt is not None:
                try:
                    if not filt(state):
                        continue
                except Exception as e:
                    # 过滤器签名写错（比如不接收 state 参数）也会走到这里。
                    # 静默 continue 会让字段「莫名其妙」丢失且毫无线索，
                    # 所以必须把原因单独记下来（不能写进 last_error——
                    # 那会被随后的成功写入重置掉）。
                    self.filter_errors[key] = str(e)
                    self.load_warnings.append(
                        f'配置字段 {key} 的过滤器异常，本次未保存该字段')
                    continue
            # tuples → lists（JSON 不认 tuple）
            if key in _TUPLE_FIELDS and isinstance(value, (list, tuple)):
                value = [list(x) if isinstance(x, (list, tuple)) else x
                         for x in value]
            settings[key] = value

        return {'version': CONFIG_VERSION, 'settings': settings}

    def save(self, state, force: bool = False) -> bool:
        """
        保存配置。force=True 忽略防抖立即写。

        返回是否真的落盘了。
        """
        if not self.enabled:
            return False

        with self._lock:

            payload = self.build_payload(state)

        if not payload['settings']:
            # 没有可保存的内容，删掉空文件而不是留 '{}'
            return self._clear_file()

        return self._atomic_write(payload)

    def _atomic_write(self, payload: Dict[str, Any]) -> bool:
        """先写临时文件再 os.replace，保证原子性。"""
        directory = os.path.dirname(self.path)
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as e:
            self.last_error = f'无法创建配置目录: {e}'
            return False

        tmp_path = None
        try:
            fd, tmp_path = tempfile.mkstemp(
                dir=directory, prefix='.config-', suffix='.tmp')
            with os.fdopen(fd, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
                f.flush()
                try:
                    os.fsync(f.fileno())   # 确保落盘，防断电
                except OSError:
                    pass
            os.replace(tmp_path, self.path)
            tmp_path = None
            self.last_error = None
            with self._lock:
                self._last_save = time.time()
            return True
        except OSError as e:
            self.last_error = f'配置保存失败: {e}'
            return False
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    def _clear_file(self) -> bool:
        """删除配置文件（配置为空时）。"""
        try:
            if os.path.exists(self.path):
                os.remove(self.path)
            return True
        except OSError as e:
            self.last_error = f'配置删除失败: {e}'
            return False

    # ══════════════════════════════════════════════════════════
    #  防抖
    # ══════════════════════════════════════════════════════════

    def mark_dirty(self):
        """
        标记为「有改动」，安排一次防抖保存（窗口内会合并为一次写盘）。

        为什么不立即写：发送日志每 800ms 写几十条，
        每次都落盘会把磁盘打满并卡住界面。

        _schedule_locked 内部有计时器判重，窗口内的重复调用不会叠加定时器。
        """
        if not self.enabled:
            return
        with self._lock:
            self._schedule_locked()

    def _schedule_locked(self):
        """安排防抖落盘。已挂起计时器时不重复安排（这是防抖的关键）。"""
        if self._timer is not None:
            return
        self._timer = threading.Timer(self.debounce_seconds, self._on_timer)
        self._timer.daemon = True
        self._timer.start()

    def _on_timer(self):
        """防抖窗口到期，执行落盘。"""
        with self._lock:
            self._timer = None
        self.flush()

    def flush(self):
        """
        立即落盘（含挂起的防抖计时器）。

        程序退出、开始发送等关键时机调用。
        """
        # 关掉挂起的计时器，避免它稍后再写一次
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
                self._timer = None
        if self._provider is not None:
            self.save(self._provider)

    def bind(self, state):
        """
        绑定 state 引用，`mark_dirty()` 时才知道往哪存。

        调用后会自动加载历史配置到 state（校验通过的部分）。
        """
        self._provider = state
        saved = self.load()
        if saved:
            for k, v in saved.items():
                if k in _TUPLE_FIELDS and isinstance(v, (list, tuple)):
                    v = [tuple(x) if isinstance(x, list) else x for x in v]
                state[k] = v
        return saved

    def delete(self) -> bool:
        """清除磁盘配置（用户点「清除配置」时用）。"""
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
                self._timer = None

        return self._clear_file()

    def info(self) -> dict:
        return {
            'path': self.path,
            'exists': os.path.isfile(self.path),
            'enabled': self.enabled,
            'last_error': self.last_error,
            'warnings': list(self.load_warnings),
            'filter_errors': dict(self.filter_errors),
        }

    # ══════════════════════════════════════════════════════════
    #  最近文件夹
    # ══════════════════════════════════════════════════════════

    MAX_RECENT = 8

    def remember_folder(self, state, folder: str):
        """
        把文件夹记入「最近使用」（新的排前面，重复的移到最前）。

        纯内存操作，落盘由外层的 mark_dirty 负责。
        """
        if not folder or not os.path.isdir(folder):
            return
        norm = os.path.normpath(folder)
        recent = list(state.get('recent_folders', []) or [])
        # 去掉同一个（Windows 路径不区分大小写）
        recent = [p for p in recent
                  if p.lower() != norm.lower()]
        recent.insert(0, norm)
        del recent[self.MAX_RECENT:]
        state['recent_folders'] = recent
