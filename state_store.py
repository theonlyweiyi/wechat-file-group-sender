#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
线程安全状态容器
================
替代原先的裸 dict 全局 `state`。

为什么要换
----------
原实现 `state` 是模块级 dict，**没有任何锁**。而实际使用中：
  - `_send_worker` 在后台线程里写 `state["logs"]` / `state["progress"]`
  - 前端每 800ms 轮询 `/api/status` 读同一份 state
这依赖 GIL 恰好不被打断。原代码里 `state["logs"].append(...)` 与
`state["logs"] = state["logs"][-200:]` 两步不是原子的，
理论上前端可能拿到超过 200 条或丢日志。多线程下这是真实的竞态。

为什么保留 dict 语法
-------------------
`web_app.py` 里有 62 处 `state["key"]` 访问。保留 `__getitem__` /
`__setitem__` 可以让这些调用点**一行都不用改**，把改动风险压到最小。

自动持久化 + 防抖
-----------------
`__setitem__` 标记「脏」并请求保存，但不立即写盘——否则发送日志每
800ms 写几十条会把磁盘打满、把界面卡死。真正的落盘由 `ConfigStore`
按时间窗口防抖，或在关键时机（切换文件夹/预览/发送前/退出）显式触发。
"""

import threading
import time
from typing import Any, Callable, Dict, List, Optional


class StateStore:
    """
    线程安全的键值状态容器，接口兼容 dict 的下标访问。

    典型用法：
        state = StateStore(initial)
        state["folder"] = "D:/x"          # 触发脏标记 + 钩子
        state.get("folder", "")           # 带默认值读取
        state.update(a=1, b=2)            # 原子批量更新
        state.atomic_update(fn)           # 在锁内执行读-改-写
    """

    def __init__(self, initial: Optional[Dict[str, Any]] = None,
                 on_change: Optional[Callable[[str, Any], None]] = None,
                 log_limit: int = 200):
        """
        initial  初始字段
        on_change 变更回调 (key, value)，用于通知持久化层。不在锁内调用。
        log_limit logs 字段的自动裁剪上限，超出则丢最旧的。
        """
        self._lock = threading.RLock()   # 可重入：atomic_update 内再 __setitem__ 不会死锁
        self._data: Dict[str, Any] = dict(initial or {})
        self._on_change = on_change
        self.log_limit = log_limit
        # 读快照计数：用于测试与调试观察竞态
        self._writes = 0
        self._reads = 0

    # ── dict 兼容接口 ──────────────────────────────────────

    def __getitem__(self, key):
        with self._lock:
            self._reads += 1
            return self._data[key]

    def __setitem__(self, key, value):
        with self._lock:
            self._writes += 1
            self._data[key] = value
            # logs 特殊处理：自动裁剪，避免无限增长
            if key == 'logs' and self.log_limit and isinstance(value, list):
                if len(value) > self.log_limit:
                    self._data[key] = value[-self.log_limit:]
        self._notify(key, self._data.get(key))

    def __delitem__(self, key):
        with self._lock:
            del self._data[key]
        self._notify(key, None)

    def __contains__(self, key):
        with self._lock:
            return key in self._data

    def __len__(self):
        with self._lock:
            return len(self._data)

    def __iter__(self):
        with self._lock:
            return iter(list(self._data.keys()))

    def __repr__(self):
        with self._lock:
            return f'StateStore({self._data!r})'

    # ── dict 常用方法 ──────────────────────────────────────

    def get(self, key, default=None):
        with self._lock:
            return self._data.get(key, default)

    def keys(self):
        with self._lock:
            return list(self._data.keys())

    def values(self):
        with self._lock:
            return list(self._data.values())

    def items(self):
        with self._lock:
            return list(self._data.items())

    def setdefault(self, key, default=None):
        with self._lock:
            if key not in self._data:
                self._data[key] = default
                self._writes += 1
                self._notify(key, default)
            return self._data[key]

    def pop(self, key, *default):
        with self._lock:
            had = key in self._data
            val = self._data.pop(key, *default) if had else (default[0] if default else None)
            if had:
                self._notify(key, None)
            return val

    def clear(self):
        with self._lock:
            self._data.clear()
        self._notify('*', None)

    # ── StateStore 特有 ────────────────────────────────────

    def update(self, _sentinel=None, **kwargs):
        """批量更新。兼容 dict.update(**kwargs)。"""
        items = {}
        if _sentinel is not None:
            items.update(_sentinel)
        items.update(kwargs)
        with self._lock:
            for k, v in items.items():
                self._data[k] = v
                self._writes += 1
        for k in items:
            self._notify(k, self._data.get(k))

    def atomic_update(self, mutator: Callable[[Dict[str, Any]], Any]):
        """
        在锁内执行「读-改-写」。

        mutator 接收一个可变的内部 dict 引用，修改后调用方无需再set。
        返回 mutator 的返回值。

        这是替代「先 get 再 set」的正确方式——后者在多线程下会丢更新：

            # 错误写法
            logs = state['logs']; logs.append(x); state['logs'] = logs

            # 正确写法
            state.atomic_update(lambda d: d['logs'].append(x))
        """
        with self._lock:
            result = mutator(self._data)
            if isinstance(self._data.get('logs'), list) and self.log_limit:
                if len(self._data['logs']) > self.log_limit:
                    self._data['logs'] = self._data['logs'][-self.log_limit:]
            return result

    def snapshot(self) -> Dict[str, Any]:
        """
        取当前状态的**深拷贝**（list/dict 逐层复制）。

        为什么要深拷贝：直接 `dict(self._data)` 是浅拷贝，
        调用方拿到后修改嵌套的 list（如 logs）会**直接影响内部状态**，
        在并发下就成了一种隐蔽的数据竞争。
        """
        import copy
        with self._lock:
            return copy.deepcopy(self._data)

    def shallow_snapshot(self) -> Dict[str, Any]:
        """浅拷贝。调用方承诺不修改嵌套结构。"""
        with self._lock:
            return dict(self._data)

    def replace(self, new_data: Dict[str, Any]):
        """整体替换（用于从磁盘加载配置）。"""
        with self._lock:
            self._data.clear()
            self._data.update(new_data or {})
            self._writes += 1
        self._notify('*', None)

    def reset_to(self, initial: Dict[str, Any]):
        """恢复为初始字段集（清掉运行时残留）。"""
        self.replace(initial)

    def set_hook(self, on_change: Optional[Callable[[str, Any], None]]):
        """替换变更回调。"""
        self._on_change = on_change

    def stats(self) -> dict:
        with self._lock:
            return {
                'fields': len(self._data),
                'writes': self._writes,
                'reads': self._reads,
            }

    # ── 内部 ───────────────────────────────────────────────

    def _notify(self, key, value):
        """在锁外调用回调，避免回调里再操作 state 造成死锁。"""
        if self._on_change:
            try:
                self._on_change(key, value)
            except Exception:
                # 钩子失败绝不能影响主流程（持久化是辅助功能）
                pass


class ListAppender:
    """
    列表的高频追加代理 —— 解决「append + 裁剪」非原子问题。

    用法：
        la = ListAppender(state, 'logs')
        la.append({"time": ..., "msg": ...})   # 原子
    """

    def __init__(self, store: StateStore, key: str):
        self._store = store
        self._key = key

    def append(self, item):
        def mutator(d):
            lst = d.setdefault(self._key, [])
            if not isinstance(lst, list):
                lst = []
                d[self._key] = lst
            lst.append(item)
            n = self._store.log_limit
            if n and len(lst) > n:
                # 原地裁剪，避免重新绑定 key
                del lst[:len(lst) - n]
        self._store.atomic_update(mutator)
        self._store._notify(self._key, item)

    def extend(self, items):
        def mutator(d):
            lst = d.setdefault(self._key, [])
            if not isinstance(lst, list):
                lst = []
                d[self._key] = lst
            lst.extend(items)
        self._store.atomic_update(mutator)
        self._store._notify(self._key, None)

    def snapshot(self) -> List:
        return self._store.get(self._key, [])

    def clear(self):
        self._store[self._key] = []
