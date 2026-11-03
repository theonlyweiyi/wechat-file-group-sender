#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信文件群发工具 — Web 版
==========================
Flask + 现代化 Web UI，浏览器打开即用。

启动后访问: http://localhost:5890
"""

import os
import sys
import time
import json
import threading
from datetime import datetime
from io import BytesIO

from flask import Flask, request, jsonify, send_from_directory, send_file, Response
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from wechat_sender import WeChatSender, set_log_fn as sender_set_log
from send_outcome import SendOutcome, SendResult, Stage
from retry import RetryPolicy, RetryQueue, EntryStatus
from state_store import StateStore, ListAppender
from config_store import ConfigStore, PERSIST_FIELDS

app = Flask(__name__)
app.secret_key = 'wechat_sender_secret_key_2026'


def _runtime_dir() -> str:
    """
    运行期可写目录。

    打包成单文件 exe 后，程序运行在 %TEMP%\\_MEIxxxxxx 里——那是随进程
    结束就消失的临时解压目录。往那儿写文件等于白写（下次启动就没了），
    而且源码运行与打包运行的路径还不一样。

    所以统一用 config_store 定好的用户配置目录，与配置文件放在一起。
    """
    from config_store import _default_config_dir
    path = _default_config_dir()
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        path = os.path.join(os.path.expanduser('~'), '.wechat_file_sender')
        os.makedirs(path, exist_ok=True)
    return path


UPLOAD_FOLDER = os.path.join(_runtime_dir(), '_uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

TEMPLATE_NAME = "发送映射模板.xlsx"

def _resolve_template_source():
    """
    定位内嵌的模板资源。

    打包后模板被编译进 exe 资源，运行时解压在 sys._MEIPASS 下——
    那是临时目录，所以这条路径只在程序运行期间有效，正好够用。
    源码运行时则从脚本目录读。
    """
    candidates = []
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, TEMPLATE_NAME))
    candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), TEMPLATE_NAME))
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None

def _build_template_bytes():
    """动态生成模板 xlsx 的字节流（兜底方案，保证单文件零附带）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "文件发送映射"

    header_font = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="07C160", end_color="07C160", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    cell_font = Font(name="Microsoft YaHei", size=10)
    cell_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    headers = ["文件名关键词", "微信备注名/昵称", "备注（可选）"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    examples = [
        ("月报", "张三", "文件名包含【月报】即匹配"),
        ("工资单", "李四", "如: 2024年08月工资单.xlsx"),
        ("合同", "王五", "如: 劳动合同-张三.pdf"),
        ("发票", "赵六", "如: 增值税发票_001.pdf"),
    ]
    for row, (keyword, name, note) in enumerate(examples, 2):
        for col, val in enumerate([keyword, name, note], 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = cell_font
            cell.alignment = cell_align
            cell.border = thin_border

    note_row = len(examples) + 3
    ws.cell(row=note_row, column=1, value="使用说明:").font = Font(name="Microsoft YaHei", size=10, bold=True)
    notes = [
        "1. 第一列填文件名关键词，第二列填对方的微信备注名（优先）或昵称",
        "2. 匹配规则: 文件名包含关键词即匹配（不区分大小写）",
        "3. 一个关键词匹配到多个文件时，全部发送给对应联系人",
        "4. 多个关键词映射到同一个人时会自动合并去重，只发送一次",
        "5. 第一行为表头，程序自动跳过",
        "6. 请确保微信PC版已登录；重名时建议填微信号（唯一）",
    ]
    for i, note in enumerate(notes, 1):
        ws.cell(row=note_row + i, column=1, value=note).font = Font(name="Microsoft YaHei", size=9, color="888888")

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 40

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── 全局状态 ───────────────────────────────────────────────
# 缺口4：原先是裸 dict，无锁、无持久化。现改为 StateStore：
#   - 保留 state["key"] 下标语法，下方 60+ 处调用点无需改动
#   - 内部 RLock 解决「_send_worker 写/ 前端轮询读」的竞态
#   - 变更时经config_store 落盘，下次启动自动恢复
_INITIAL_STATE = {
    "folder_path": "",
    "excel_uploaded": None,        # (filename, path) —— 运行时，不持久化
    "mappings": [],                 # [(keyword, name), ...]
    "matched": [],                  # 明细: [{keyword, name, files}, ...]
    "grouped": [],                  # 发送计划(按联系人聚合去重): [{name, keywords, files}, ...]
    "unmatched_files": [],          # 文件夹里没匹配到任何人的文件
    "sending": False,
    "stop_requested": False,
    "logs": [],
    "progress": {"current": 0, "total": 0},
    "result": None,                 # "completed" / "stopped" / None
    # ── 失败重试（缺口2）──
    "max_retries": 2,                # 每个联系人最多尝试次数（含首次）
    "queue": [],                     # 全部发送条目（含每次尝试的历史）
    "needs_review": [],              # ⚠ 状态不明、等人工判断的条目
    "stats": None,                   # 本轮统计
    "active_queue": None,            # 运行中的 RetryQueue 实例（不持久化）
    # ── 持久化相关（缺口4）──
    "enable_msg": False,             # 是否附带消息
    "msg": "",                       # 附言模板
    "recent_folders": [],            # 最近使用文件夹
    "window_geometry": None,         # 窗口位置尺寸
    "config_restored": False,        # 本次启动是否从配置恢复（供前端提示）
    "config_warnings": [],           # 加载配置时的提示
}

# 配置存储（%APPDATA%\WeChatFileSender\config.json）
config_store = ConfigStore(debounce_seconds=2.0)
INITIAL_STATE_SNAPSHOT = dict(_INITIAL_STATE)

# 只有「可持久化字段」的变化才触发写盘，避免日志高频写拖累性能
_PERSIST_SET = set(PERSIST_FIELDS)


def _on_state_change(key, value):
    """StateStore 变更钩子：仅对白名单字段触发防抖保存。"""
    if key in _PERSIST_SET or key == '*':
        config_store.mark_dirty()


# 哪类来源的映射会持久化。
# 只有「明确来自上传的Excel」才排除——那份文件躺在 _uploads 临时目录，
# 重启后可能已被清理，存下来下次启动就是一条失效路径。
# 其余情况（手动输入、或还没设置来源）都应保存，否则用户辛苦填的映射会丢。
#
# 注意不能用 `if src: return src[0]=='手动输入'` 这种写法：
# None（尚未选择来源）会被误判成「不是手动输入」而丢掉映射。
# 注意：ConfigStore 会以 predicate(state) 的形式调用它，必须接受一个参数。
def _mappings_are_persistable(_state=None) -> bool:
    src = state.get("excel_uploaded")
    if not src:
        return True                    # 没来源标记 → 视为可持久化
    return src[0] == "手动输入"


state = StateStore(_INITIAL_STATE, on_change=_on_state_change, log_limit=200)
config_store.add_field_filter('mappings', _mappings_are_persistable)

# 启动即加载历史配置（把校验通过的值写回 state）
_saved_settings = config_store.bind(state)
if _saved_settings:
    state["config_restored"] = True
    state["config_warnings"] = list(config_store.load_warnings)
    # 恢复的映射属于「手动输入」来源，标记以便后续继续持久化
    if state.get("mappings"):
        state["excel_uploaded"] = ("手动输入", None)

# ── 日志 ──────────────────────────────────────────────────
# 用ListAppender 做「追加+裁剪」的原子操作：_send_worker 在后台线程写，
# 前端每 800ms 轮询读，原来两步非原子会有竞态。
_log_appender = ListAppender(state, "logs")


def _add_log(msg):
    """向日志列表添加条目（线程安全）"""
    ts = datetime.now().strftime("%H:%M:%S")
    _log_appender.append({"time": ts, "msg": msg})


# ══════════════════════════════════════════════════════════════
#  前端页面
# ══════════════════════════════════════════════════════════════

HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>微信文件群发工具</title>
<link rel="icon" href="data:,">
<style>
  :root {
    --bg: #f0f2f5;
    --card: #fff;
    --primary: #07c160;
    --primary-hover: #06ad56;
    --danger: #e74c3c;
    --text: #1f2937;
    --text-light: #6b7280;
    --text-muted: #9ca3af;
    --border: #e5e7eb;
    --input-bg: #f9fafb;
    --shadow: 0 1px 3px rgba(0,0,0,.08), 0 1px 2px rgba(0,0,0,.06);
    --radius: 10px;
    --radius-sm: 6px;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
  }

  .container { max-width: 860px; margin: 0 auto; padding: 24px 20px; }

  /* Header */
  .header { text-align: center; padding: 32px 0; }
  .header h1 { font-size: 26px; font-weight: 700; color: var(--text); }
  .header p { color: var(--text-light); margin-top: 6px; font-size: 14px; }

  /* Card */
  .card {
    background: var(--card);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    padding: 24px;
    margin-bottom: 16px;
  }
  .card-title {
    font-size: 15px; font-weight: 600; color: var(--text);
    margin-bottom: 16px; display: flex; align-items: center; gap: 8px;
  }
  .card-title .icon { font-size: 18px; }

  /* Form */
  .form-row { display: flex; gap: 10px; align-items: flex-end; flex-wrap: wrap; }
  .form-group { flex: 1; min-width: 200px; }
  .form-group label { display: block; font-size: 13px; font-weight: 600; color: var(--text-light); margin-bottom: 6px; }
  .form-group input, .form-group textarea {
    width: 100%; padding: 10px 14px; font-size: 14px;
    border: 1.5px solid var(--border); border-radius: var(--radius-sm);
    background: var(--input-bg); color: var(--text);
    transition: border-color .15s; outline: none; font-family: inherit;
  }
  .form-group input:focus, .form-group textarea:focus { border-color: var(--primary); box-shadow: 0 0 0 3px rgba(7,193,96,.1); }
  .form-group textarea { resize: vertical; min-height: 56px; }
  .form-group .hint { font-size: 12px; color: var(--text-muted); margin-top: 4px; }
  /* 步骤4 的安全说明：加粗部分需更醒目 */
  .card .hint b { color: var(--text); font-weight: 600; }

  /* Buttons */
  .btn {
    display: inline-flex; align-items: center; gap: 6px;
    padding: 10px 20px; font-size: 14px; font-weight: 600;
    border: none; border-radius: var(--radius-sm); cursor: pointer;
    transition: all .15s; font-family: inherit;
    white-space: nowrap;
  }
  .btn:disabled { opacity: .5; cursor: not-allowed; }
  .btn-primary { background: var(--primary); color: #fff; }
  .btn-primary:hover:not(:disabled) { background: var(--primary-hover); }
  .btn-outline { background: #fff; color: var(--text); border: 1.5px solid var(--border); }
  .btn-outline:hover:not(:disabled) { background: #f9fafb; border-color: var(--primary); color: var(--primary); }
  .btn-danger { background: var(--danger); color: #fff; }
  .btn-danger:hover:not(:disabled) { background: #c0392b; }
  .btn-sm { padding: 6px 14px; font-size: 13px; }
  .btn-block { width: 100%; justify-content: center; }

  /* Badge */
  .badge {
    display: inline-block; padding: 2px 10px; font-size: 12px; font-weight: 600;
    border-radius: 20px; line-height: 1.6;
  }
  .badge-success { background: #d1fae5; color: #065f46; }
  .badge-warn { background: #fef3c7; color: #92400e; }
  .badge-fail { background: #fee2e2; color: #991b1b; }
  .badge-info { background: #dbeafe; color: #1e40af; }

  /* File list */
  .file-list {
    max-height: 180px; overflow-y: auto; margin-top: 8px;
    font-size: 13px; color: var(--text-light);
    padding: 8px 12px; background: var(--input-bg);
    border-radius: var(--radius-sm); border: 1px solid var(--border);
  }
  .file-list .file-item { padding: 3px 0; display: flex; align-items: center; gap: 6px; }
  .file-list .file-icon { color: var(--primary); }

  /* Table */
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; padding: 10px 12px; font-weight: 600; color: var(--text-light); border-bottom: 2px solid var(--border); font-size: 12px; text-transform: uppercase; letter-spacing: .5px; }
  td { padding: 10px 12px; border-bottom: 1px solid var(--border); vertical-align: top; }
  tr:hover td { background: #f9fafb; }
  .file-list-inline { font-size: 12px; color: var(--text-light); }
  .file-list-inline div { padding: 1px 0; }
  .no-data { text-align: center; color: var(--text-muted); padding: 40px 0; font-size: 14px; }

  /* Tabs */
  .tabs { display: flex; gap: 0; margin-bottom: 16px; border-bottom: 2px solid var(--border); }
  .tab {
    padding: 8px 20px; font-size: 13px; font-weight: 600; cursor: pointer;
    border: none; background: none; color: var(--text-muted);
    border-bottom: 2px solid transparent; margin-bottom: -2px;
    transition: all .15s; font-family: inherit;
  }
  .tab.active { color: var(--primary); border-bottom-color: var(--primary); }
  .tab:hover { color: var(--text); }

  /* Manual input table */
  .mapping-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  .mapping-table th { padding: 8px 10px; text-align: left; font-size: 12px; }
  .mapping-table td { padding: 4px 6px; border-bottom: 1px solid var(--border); }
  .mapping-table input {
    width: 100%; padding: 8px 10px; font-size: 13px;
    border: 1.5px solid var(--border); border-radius: var(--radius-sm);
    background: var(--input-bg); color: var(--text); outline: none;
    font-family: inherit;
  }
  .mapping-table input:focus { border-color: var(--primary); }
  .btn-icon {
    background: none; border: none; cursor: pointer; color: var(--text-muted);
    font-size: 16px; padding: 4px 8px; border-radius: var(--radius-sm);
    transition: all .15s;
  }
  .btn-icon:hover { color: var(--danger); background: #fee2e2; }
  .add-row-btn { margin-top: 10px; }

  /* Logs */
  .log-area {
    background: #1e1e1e; color: #d4d4d4; border-radius: var(--radius-sm);
    padding: 12px 16px; font-family: "Cascadia Code", "Fira Code", Consolas, monospace;
    font-size: 12px; line-height: 1.7; max-height: 260px; overflow-y: auto;
  }
  .log-line { white-space: pre-wrap; word-break: break-all; }
  .log-line.success { color: #4ecb71; }
  .log-line.fail { color: #f14c4c; }
  .log-line.info { color: #569cd6; }
  .log-line.warn { color: #cca700; }
  .log-line.review { color: #f59e0b; }
  .log-line.summary { color: #4ecb71; font-weight: 600; }

  /* 待人工确认队列 */
  .review-item {
    display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
    padding: 12px 14px; margin-bottom: 8px;
    background: #fffbeb; border: 1px solid #fde68a; border-radius: var(--radius-sm);
  }
  .review-info { flex: 1; min-width: 200px; }
  .review-name { font-size: 14px; font-weight: 600; color: var(--text); }
  .review-files { font-size: 12px; color: var(--text-light); margin-top: 2px;
    word-break: break-all; }
  .review-detail { font-size: 11px; color: var(--text-muted); margin-top: 3px; }
  .review-actions { display: flex; gap: 8px; flex-shrink: 0; }

  /* 发送汇总 */
  .stat-box {
    flex: 1; min-width: 96px; padding: 12px 10px; text-align: center;
    background: var(--input-bg); border: 1px solid var(--border);
    border-radius: var(--radius-sm);
  }
  .stat-value { font-size: 22px; font-weight: 700; line-height: 1.2; }
  .stat-label { font-size: 11px; color: var(--text-light); margin-top: 3px; }
  .stat-total .stat-value { color: var(--text); }
  .stat-ok .stat-value { color: #07c160; }
  .stat-retry .stat-value { color: #3b82f6; }
  .stat-review .stat-value { color: #f59e0b; }
  .stat-fail .stat-value { color: #e74c3c; }

  /* 汇总结论 */
  .stats-conclusion {
    margin-top: 12px; padding: 9px 12px; font-size: 13px; font-weight: 500;
    border-radius: var(--radius-sm); background: var(--input-bg);
    border: 1px solid var(--border); color: var(--text);
  }
  .stats-conclusion.ok { background: #f0fdf4; border-color: #bbf7d0; color: #15803d; }
  .stats-conclusion.warn { background: #fffbeb; border-color: #fde68a; color: #92400e; }

  /* Progress */
  .progress-wrap { margin: 14px 0; }
  .progress-bar {
    width: 100%; height: 8px; background: var(--border);
    border-radius: 4px; overflow: hidden;
  }
  .progress-fill {
    height: 100%; background: var(--primary);
    border-radius: 4px; transition: width .3s;
    width: 0%;
  }
  .progress-label {
    display: flex; justify-content: space-between;
    font-size: 12px; color: var(--text-light); margin-top: 6px;
  }

  /* Upload zone */
  .upload-zone {
    border: 2px dashed var(--border); border-radius: var(--radius-sm);
    padding: 28px; text-align: center; cursor: pointer;
    transition: all .15s; color: var(--text-muted);
  }
  .upload-zone:hover, .upload-zone.drag-over {
    border-color: var(--primary); color: var(--primary); background: rgba(7,193,96,.03);
  }
  .upload-zone .upload-icon { font-size: 32px; margin-bottom: 8px; }
  .upload-zone .upload-text { font-size: 14px; }
  .upload-zone .upload-hint { font-size: 12px; margin-top: 4px; }
  .upload-zone input[type="file"] { display: none; }

  .template-link { margin-top: 10px; text-align: center; }
  .btn-text {
    background: none; border: none; cursor: pointer;
    color: var(--brand, #07c160); font-size: 13px; padding: 4px 6px;
    text-decoration: underline; font-family: inherit;
  }
  .btn-text:hover { color: #05984c; }

  .file-uploaded {
    display: flex; align-items: center; gap: 8px;
    padding: 8px 14px; background: #d1fae5; border-radius: var(--radius-sm);
    font-size: 13px; color: #065f46; margin-top: 8px;
  }

  /* Toggle */
  .toggle-row { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
  .toggle-row label { font-size: 13px; font-weight: 600; color: var(--text-light); cursor: pointer; }

  /* Toast */
  .toast {
    position: fixed; top: 20px; left: 50%; transform: translateX(-50%);
    padding: 12px 24px; border-radius: var(--radius-sm);
    font-size: 14px; font-weight: 600; color: #fff; z-index: 9999;
    animation: toastIn .3s ease;
  }
  .toast.success { background: var(--primary); }
  .toast.error { background: var(--danger); }
  @keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(-10px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }

  /* 文件分配统计条 */
  .alloc-bar {
    padding: 12px 14px; margin-bottom: 14px;
    background: var(--input-bg); border: 1px solid var(--border);
    border-radius: var(--radius-sm);
  }
  .alloc-numbers { display: flex; align-items: center; gap: 16px; flex-wrap: wrap; }
  .alloc-stat { text-align: center; min-width: 72px; }
  .alloc-value { font-size: 21px; font-weight: 700; line-height: 1.15; color: var(--text); }
  .alloc-value.ok { color: var(--primary); }
  .alloc-value.danger { color: var(--danger); }
  .alloc-label { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
  .alloc-op { font-size: 17px; color: var(--text-muted); }
  .alloc-coverage { margin-left: auto; min-width: 150px; }
  .cov-label { font-size: 11px; color: var(--text-light); margin-bottom: 4px; }
  .cov-label b { color: var(--text); font-size: 13px; }
  .cov-track { height: 7px; background: var(--border); border-radius: 4px; overflow: hidden; }
  .cov-fill { height: 100%; width: 0; background: var(--primary); border-radius: 4px;
    transition: width .3s ease; }
  .cov-fill.warn { background: #f59e0b; }
  .cov-fill.danger { background: var(--danger); }

  /* 未匹配文件清单 */
  .unmatched-box {
    margin-top: 12px; padding: 9px 12px;
    background: #fef2f2; border: 1px solid #fecaca; border-radius: var(--radius-sm);
  }
  .um-title {
    font-size: 12px; color: #991b1b; cursor: pointer;
    display: flex; align-items: center; gap: 8px; user-select: none;
  }
  .um-title b { color: #7f1d1d; }
  .um-toggle { margin-left: auto; font-size: 11px; opacity: .75; }
  .unmatched-list {
    margin-top: 8px; padding-top: 8px; border-top: 1px dashed #fca5a5;
    display: flex; flex-wrap: wrap; gap: 4px; max-height: 160px; overflow-y: auto;
  }
  .unmatched-item {
    font-size: 11px; padding: 2px 7px; border-radius: 9px;
    background: #fff; border: 1px solid #fca5a5; color: #991b1b;
    word-break: break-all;
  }
  /* 文件多时的「还有 N 个」提示 */
  .unmatched-more {
    font-size: 11px; color: #991b1b; align-self: center;
    padding: 2px 6px; opacity: .8;
  }

  /* 预览视图切换 */
  .view-switch { display: flex; align-items: center; gap: 6px; margin-bottom: 12px; flex-wrap: wrap; }  .vs-btn {
    padding: 5px 12px; font-size: 12px; border-radius: 14px; cursor: pointer;
    border: 1px solid var(--border); background: #fff; color: var(--text-light);
    font-family: inherit; transition: all .15s;
  }
  .vs-btn:hover { border-color: var(--primary); color: var(--primary); }
  .vs-btn.active { background: var(--primary); border-color: var(--primary); color: #fff; }
  .vs-tip { margin-left: auto; font-size: 12px; color: var(--text-muted); }
  .vs-tip b { color: #d97706; }
  .kw-tags { display: flex; flex-wrap: wrap; gap: 4px; }
  .kw-tag {
    font-size: 11px; padding: 1px 7px; border-radius: 10px;
    background: #f1f5f9; color: #475569; border: 1px solid var(--border);
  }

  /* 配置恢复提示条 */
  .restore-banner {
    display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
    padding: 10px 14px; margin-bottom: 14px; font-size: 13px;
    background: #eff6ff; border: 1px solid #bfdbfe; border-radius: var(--radius-sm);
    color: #1e40af;
  }
  .restore-banner .rb-actions { margin-left: auto; display: flex; gap: 6px; }
  .restore-banner.warn { background: #fffbeb; border-color: #fde68a; color: #92400e; }

  /* 最近文件夹下拉 */
  .recent-row { display: flex; align-items: center; gap: 8px; margin-top: 8px; flex-wrap: wrap; }
  .recent-row select { flex: 1; min-width: 0; font-size: 12px; padding: 6px 10px; }
  .recent-label { font-size: 12px; color: var(--text-muted); flex-shrink: 0; }

  /* 配置路径提示 */
  .cfg-path { font-size: 11px; color: var(--text-muted); word-break: break-all; }

  /* Responsive */
  @media (max-width: 640px) {
    .container { padding: 12px; }
    .card { padding: 16px; }
    .form-row { flex-direction: column; }
  }
</style>
</head>
<body>
<div class="container">

  <div class="header">
    <h1>📁 微信文件群发工具</h1>
    <p>按关键词匹配文件，自动发送给对应的微信联系人</p>
  </div>

  <!-- 配置恢复提示（启动时若读到已保存的配置） -->
  <div class="restore-banner" id="restoreBanner" style="display:none;">
    <span id="restoreText">已恢复上次使用的设置</span>
    <div class="rb-actions">
      <button class="btn btn-sm btn-outline" onclick="clearConfig()">清除配置</button>
      <button class="btn btn-sm btn-outline" onclick="hideRestoreBanner()">知道了</button>
    </div>
  </div>

  <!-- 步骤 1: 文件夹 -->
  <div class="card">
    <div class="card-title"><span class="icon">📂</span> 步骤 1：选择文件所在的文件夹</div>
    <div class="form-row">
      <div class="form-group">
        <label>文件夹路径</label>
        <input type="text" id="folderPath" placeholder="例如: D:\工作\8月文件">
      </div>
      <div style="padding-bottom: 2px; display: flex; gap: 8px;">
        <button class="btn btn-outline" onclick="selectFolder()">📂 浏览选择</button>
        <button class="btn btn-outline" onclick="scanFolder()" id="btnScan">🔍 检查文件夹</button>
      </div>
    </div>
    <div class="recent-row" id="recentRow" style="display:none;">
      <span class="recent-label">最近使用：</span>
      <select id="recentFolders" onchange="applyRecentFolder(this.value)">
        <option value="">— 选择历史文件夹 —</option>
      </select>
    </div>
    <div id="folderFiles" class="file-list" style="display:none;"></div>
    <div class="cfg-path" id="configPathHint" style="display:none;margin-top:8px;"></div>
  </div>

  <!-- 步骤 2: 映射表 -->
  <div class="card">
    <div class="card-title"><span class="icon">📋</span> 步骤 2：配置映射关系</div>
    <div class="tabs">
      <button class="tab active" onclick="switchMappingTab('upload')" id="tabUpload">📄 上传 Excel</button>
      <button class="tab" onclick="switchMappingTab('manual')" id="tabManual">✏️ 手动输入</button>
    </div>
    <!-- 上传模式 -->
    <div id="sectionUpload">
      <div class="upload-zone" id="uploadZone" onclick="document.getElementById('excelFile').click()">
        <div class="upload-icon">📄</div>
        <div class="upload-text">点击选择或拖拽 Excel 文件到此处</div>
        <div class="upload-hint">支持 .xlsx / .xls / .csv 格式 | 第一行为表头，自动跳过</div>
        <input type="file" id="excelFile" accept=".xlsx,.xls,.csv" onchange="handleExcelUpload(event)">
      </div>
      <div class="template-link">
        <button class="btn-text" onclick="downloadTemplate()">⬇ 没有模板？点此下载映射表模板</button>
      </div>
      <div id="excelStatus" style="display:none;"></div>
    </div>
    <!-- 手动输入模式 -->
    <div id="sectionManual" style="display:none;">
      <div class="table-wrap">
        <table class="mapping-table" id="mappingTable">
          <thead><tr><th style="width:45%">文件名关键词</th><th style="width:45%">微信备注名 / 昵称</th><th style="width:50px;"></th></tr></thead>
          <tbody id="mappingBody">
            <tr>
              <td><input type="text" placeholder="例如: 月报" class="kw-input"></td>
              <td><input type="text" placeholder="例如: 张三（建议填备注名）" class="name-input"></td>
              <td><button class="btn-icon" onclick="removeMappingRow(this)" title="删除">✕</button></td>
            </tr>
          </tbody>
        </table>
      </div>
      <div class="hint" style="margin-top:8px;">
        第二列请填对方的<b>微信备注名</b>（最准，微信搜索权重高）或<b>昵称</b>；重名时填<b>微信号</b>最保险。
        多个关键词映射同一个人时会自动合并，<b>只发送一次</b>，不会重复发。
      </div>
      <button class="btn btn-outline btn-sm add-row-btn" onclick="addMappingRow()">➕ 添加一行</button>
      <button class="btn btn-primary btn-sm add-row-btn" onclick="saveManualMappings()" style="margin-left:8px;">💾 保存映射</button>
      <div id="manualStatus" style="display:none;margin-top:8px;"></div>
    </div>
  </div>

  <!-- 步骤 3: 消息 -->
  <div class="card">
    <div class="card-title"><span class="icon">💬</span> 步骤 3：附带消息（可选）</div>
    <div class="toggle-row">
      <input type="checkbox" id="enableMsg" onclick="toggleMessage(); persistSettings()">
      <label for="enableMsg">发送文件时附带一条消息</label>
    </div>
    <div class="form-group" id="msgGroup" style="display:none;">
      <textarea id="msgText" placeholder="输入消息内容，支持变量: {name}=备注名/昵称"
        oninput="persistSettings()"></textarea>
      <div class="hint">提示: 使用 <code>{name}</code> 会自动替换为对方的备注名/昵称</div>
    </div>
  </div>

  <!-- 步骤 4: 失败重试 -->
  <div class="card">
    <div class="card-title"><span class="icon">🔁</span> 步骤 4：失败重试</div>
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
      <label for="maxRetries" style="font-size:13px;color:var(--text-light);">失败自动重试次数：</label>
      <select id="maxRetries" style="width:auto;min-width:80px;" onchange="persistSettings()">
        <option value="1">不重试</option>
        <option value="2" selected>2 次（推荐）</option>
        <option value="3">3 次</option>
        <option value="4">4 次</option>
        <option value="5">5 次</option>
      </select>
      <span class="hint" style="margin:0;">含首次发送，按 1s → 3s → 9s 递增等待</span>
    </div>
    <div class="hint" style="margin-top:10px;line-height:1.7;">
      仅对<b>确定没发出去</b>的失败自动重试。若回车已按下但无法确认结果（可能文件已发出），
      系统<b>不会自动重发</b>，而是列入下方「待人工确认」，避免给对方发两份文件。
    </div>
  </div>

  <!-- 操作按钮 -->
  <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;">
    <button class="btn btn-primary" onclick="previewMatches()" id="btnPreview">🔍 预览匹配结果</button>
    <button class="btn btn-primary" onclick="startSend()" id="btnSend" disabled>🚀 开始发送</button>
    <button class="btn btn-danger" onclick="stopSend()" id="btnStop" disabled>⏹ 停止</button>
    <span id="sendStatus" style="font-size:13px;color:var(--text-light);display:flex;align-items:center;margin-left:auto;"></span>
  </div>

  <!-- 匹配预览 -->
  <div class="card" id="previewCard" style="display:none;">
    <div class="card-title"><span class="icon">📊</span> 匹配预览</div>

    <!-- 文件分配统计（回答「是不是所有文件都找到人了」）-->
    <div class="alloc-bar" id="allocBar">
      <div class="alloc-numbers">
        <div class="alloc-stat">
          <div class="alloc-value" id="allocTotal">0</div>
          <div class="alloc-label">文件夹文件总数</div>
        </div>
        <div class="alloc-op">/</div>
        <div class="alloc-stat">
          <div class="alloc-value ok" id="allocAssigned">0</div>
          <div class="alloc-label">已分配到人</div>
        </div>
        <div class="alloc-stat" id="allocUnmatchedWrap" style="display:none;">
          <div class="alloc-value danger" id="allocUnmatched">0</div>
          <div class="alloc-label">未匹配（不会发送）</div>
        </div>
        <div class="alloc-stat">
          <div class="alloc-value" id="allocContacts">0</div>
          <div class="alloc-label">收件人数</div>
        </div>
        <div class="alloc-coverage">
          <div class="cov-label">覆盖率 <b id="allocCoverage">0%</b></div>
          <div class="cov-track"><div class="cov-fill" id="allocCovFill"></div></div>
        </div>
      </div>
      <!-- 未匹配文件清单：让被漏掉的文件可见（默认展开）-->
      <div class="unmatched-box" id="unmatchedBox" style="display:none;">
        <div class="um-title" onclick="toggleUnmatched()">
          <span>⚠ 以下文件没有匹配到任何联系人，<b>不会被发送</b>：</span>
          <span class="um-toggle" id="umToggle">收起 ▴</span>
        </div>
        <div class="unmatched-list" id="unmatchedList" style="display:flex;"></div>
      </div>
    </div>

    <div class="view-switch">
      <button class="vs-btn active" id="vsContact" onclick="switchPreviewView('contact')">按联系人（实际发送计划）</button>
      <button class="vs-btn" id="vsKeyword" onclick="switchPreviewView('keyword')">按关键词（明细）</button>
      <span class="vs-tip" id="mergeTip"></span>
    </div>
    <div class="table-wrap">
      <table>
        <thead id="previewHead"></thead>
        <tbody id="previewBody"></tbody>
      </table>
    </div>
    <div class="no-data" id="previewEmpty">暂无匹配数据，请先点击「预览匹配结果」</div>
  </div>

  <!-- 进度 -->
  <div id="progressArea" style="display:none;">
    <div class="progress-wrap">
      <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
      <div class="progress-label">
        <span id="progressText">0 / 0</span>
        <span id="progressPercent">0%</span>
      </div>
    </div>
  </div>

  <!-- 待人工确认（回车已按下但无法确认结果） -->
  <div class="card" id="reviewCard" style="display:none;border-left:3px solid #f59e0b;">
    <div class="card-title"><span class="icon">❓</span> 待人工确认
      <span style="font-size:12px;font-weight:400;color:var(--text-light);margin-left:8px;">
        这些联系人的回车键已按下，文件可能已发出。为避免重复发送，系统没有自动重试。
      </span>
    </div>
    <div id="reviewList"></div>
  </div>

  <!-- 发送汇总 -->
  <div class="card" id="statsCard" style="display:none;">
    <div class="card-title"><span class="icon">📈</span> 发送汇总</div>
    <div id="statsBody" style="display:flex;gap:10px;flex-wrap:wrap;"></div>
    <div class="stats-conclusion" id="statsConclusion" style="display:none;"></div>
  </div>

  <!-- 日志 -->
  <div class="card">
    <div class="card-title"><span class="icon">📝</span> 发送日志</div>
    <div class="log-area" id="logArea">
      <div class="log-line info">就绪，等待操作...</div>
    </div>
  </div>

</div>

<script>
// ══════════════════════════════════════════════════════════════
//  设置加载与持久化（缺口4）
// ══════════════════════════════════════════════════════════════
//  后端已把配置写到 %APPDATA%\WeChatFileSender\config.json，
//  这里负责启动时读回来填进界面，并在用户改动后写回。

async function loadSettings() {
  let s;
  try {
    const res = await fetch('/api/settings');
    s = await res.json();
  } catch (e) {
    return;                // 读不到就按默认走，不打扰用户
  }

  // 文件夹
  if (s.folder_path) {
    document.getElementById('folderPath').value = s.folder_path;
  }

  // 附言
  document.getElementById('enableMsg').checked = !!s.enable_msg;
  if (s.msg) document.getElementById('msgText').value = s.msg;
  toggleMessage();

  // 重试次数
  if (s.max_retries) {
    const sel = document.getElementById('maxRetries');
    if ([...sel.options].some(o => o.value === String(s.max_retries))) {
      sel.value = String(s.max_retries);
    }
  }

  // 映射表 → 恢复到「手动输入」页签
  if (s.mappings && s.mappings.length) {
    switchMappingTab('manual');
    const body = document.getElementById('mappingBody');
    body.innerHTML = '';
    s.mappings.forEach(([kw, nm]) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><input type="text" class="kw-input" value="${escapeAttr(kw)}"></td>
        <td><input type="text" class="name-input" value="${escapeAttr(nm)}"></td>
        <td><button class="btn-icon" onclick="removeMappingRow(this)" title="删除">✕</button></td>`;
      body.appendChild(tr);
    });
    const st = document.getElementById('manualStatus');
    st.style.display = 'block';
    st.innerHTML = `<div class="file-uploaded">📝 已恢复 ${s.mappings.length} 条映射</div>`;
  }

  // 最近文件夹
  renderRecentFolders(s.recent_folders || []);

  // 配置路径提示
  if (s.config_path) {
    const hint = document.getElementById('configPathHint');
    hint.textContent = '设置自动保存到：' + s.config_path;
    hint.style.display = 'block';
  }

  // 恢复提示条
  if (s.restored) {
    const warn = (s.warnings || []).length > 0;
    const banner = document.getElementById('restoreBanner');
    banner.className = 'restore-banner' + (warn ? ' warn' : '');
    document.getElementById('restoreText').textContent = warn
      ? '部分设置已恢复：' + s.warnings.join('；')
      : '已恢复上次使用的文件夹、映射关系与选项';
    banner.style.display = 'flex';
  }
}

function renderRecentFolders(list) {
  const row = document.getElementById('recentRow');
  const sel = document.getElementById('recentFolders');
  if (!list || !list.length) { row.style.display = 'none'; return; }
  sel.innerHTML = '<option value="">— 选择历史文件夹 —</option>'
    + list.map(p => `<option value="${escapeAttr(p)}">${escapeHtml(p)}</option>`).join('');
  row.style.display = 'flex';
}

async function applyRecentFolder(path) {
  if (!path) return;
  document.getElementById('folderPath').value = path;
  await scanFolder();
}

function hideRestoreBanner() {
  document.getElementById('restoreBanner').style.display = 'none';
}

async function clearConfig() {
  if (!confirm('确定清除已保存的配置？\n\n将同时重置当前界面上的文件夹、映射关系和选项。')) return;
  const res = await fetch('/api/settings', { method: 'DELETE' });
  const data = await res.json();
  if (data.error) return toast(data.error, 'error');

  // 重置界面
  document.getElementById('folderPath').value = '';
  document.getElementById('msgText').value = '';
  document.getElementById('enableMsg').checked = false;
  document.getElementById('maxRetries').value = '2';
  document.getElementById('mappingBody').innerHTML = '';
  document.getElementById('folderFiles').style.display = 'none';
  document.getElementById('previewCard').style.display = 'none';
  document.getElementById('recentRow').style.display = 'none';
  document.getElementById('configPathHint').style.display = 'none';
  toggleMessage();
  hideRestoreBanner();
  toast('已清除保存的配置', 'success');
}

// 把界面上的附言/重试设置存回后端（防抖）
let settingsSaveTimer = null;
function persistSettings() {
  clearTimeout(settingsSaveTimer);
  settingsSaveTimer = setTimeout(async () => {
    try {
      await fetch('/api/settings', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({
          enable_msg: document.getElementById('enableMsg').checked,
          msg: document.getElementById('msgText').value,
          max_retries: document.getElementById('maxRetries').value,
        })
      });
    } catch (e) { /* 存不上也不该打扰用户 */ }
  }, 600);
}

// ─── 映射表模式切换 ──────────────────────────────────────────
let mappingMode = 'upload';  // 'upload' | 'manual'

async function switchMappingTab(mode) {
  mappingMode = mode;
  document.getElementById('sectionUpload').style.display = mode === 'upload' ? 'block' : 'none';
  document.getElementById('sectionManual').style.display = mode === 'manual' ? 'block' : 'none';
  document.getElementById('tabUpload').classList.toggle('active', mode === 'upload');
  document.getElementById('tabManual').classList.toggle('active', mode === 'manual');
}

function addMappingRow() {
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" placeholder="例如: 月报" class="kw-input"></td>
    <td><input type="text" placeholder="例如: 张三" class="name-input"></td>
    <td><button class="btn-icon" onclick="removeMappingRow(this)" title="删除">✕</button></td>`;
  document.getElementById('mappingBody').appendChild(tr);
}

function removeMappingRow(btn) {
  const tbody = document.getElementById('mappingBody');
  if (tbody.children.length <= 1) return; // 至少保留一行
  btn.closest('tr').remove();
}

async function saveManualMappings() {
  const rows = document.querySelectorAll('#mappingBody tr');
  const mappings = [];
  for (const row of rows) {
    const kw = row.querySelector('.kw-input').value.trim();
    const name = row.querySelector('.name-input').value.trim();
    if (kw && name) mappings.push([kw, name]);
  }
  if (mappings.length === 0) return toast('请至少填写一行关键词和联系人（备注名/昵称）', 'error');

  const res = await fetch('/api/set_mappings', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({mappings})
  });
  const data = await res.json();
  const el = document.getElementById('manualStatus');
  el.style.display = 'block';
  el.innerHTML = `<div class="file-uploaded">📝 已保存 — ${data.count} 条映射</div>`;
  document.getElementById('excelStatus').style.display = 'none';
  toast(`已保存 ${data.count} 条映射`, 'success');
}
async function selectFolder() {
  // 优先使用 pywebview 的原生文件夹选择对话框（桌面模式）
  if (window.pywebview && window.pywebview.api && window.pywebview.api.select_folder) {
    try {
      const path = await window.pywebview.api.select_folder();
      if (path) {
        document.getElementById('folderPath').value = path;
        scanFolder();
      }
      return;
    } catch (e) { /* 走兜底 */ }
  }
  // 兜底：后端弹对话框
  const res = await fetch('/api/select_folder', { method: 'POST' });
  const data = await res.json();
  if (data.path) {
    document.getElementById('folderPath').value = data.path;
    scanFolder();
  }
}

// ─── 文件夹扫描 ────────────────────────────────────────────
async function scanFolder() {
  const path = document.getElementById('folderPath').value.trim();
  if (!path) return toast('请先输入文件夹路径', 'error');

  const res = await fetch('/api/scan_folder', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({path})
  });
  const data = await res.json();
  const el = document.getElementById('folderFiles');

  if (data.error) {
    el.style.display = 'block';
    el.innerHTML = `<span style="color:var(--danger);">⚠ ${escapeHtml(data.error)}</span>`;
    return;
  }

  el.style.display = 'block';
  if (data.files.length === 0) {
    el.innerHTML = '<span style="color:var(--text-muted);">📭 文件夹为空</span>';
  } else {
    el.innerHTML = `<div style="margin-bottom:4px;font-weight:600;">找到 ${data.count} 个文件:</div>` +
      data.files.map(f =>
        `<div class="file-item"><span class="file-icon">📄</span> ${escapeHtml(f)}</div>`
      ).join('');
  }
}

// ─── Excel 上传 ────────────────────────────────────────────
async function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch('/api/upload_excel', { method: 'POST', body: formData });
  const data = await res.json();
  const el = document.getElementById('excelStatus');

  if (data.error) {
    el.style.display = 'block';
    el.innerHTML = `<div style="color:var(--danger);margin-top:8px;">⚠ ${escapeHtml(data.error)}</div>`;
  } else {
    el.style.display = 'block';
    el.innerHTML = `<div class="file-uploaded">📑 ${escapeHtml(data.filename)} — ${data.count} 条映射</div>`;
  }
}

// ─── 下载模板 ──────────────────────────────────────────────
async function downloadTemplate() {
  // 桌面模式：走 pywebview 原生「另存为」对话框（WebView2 不处理 <a download>）
  if (window.pywebview && window.pywebview.api && window.pywebview.api.save_template) {
    try {
      const path = await window.pywebview.api.save_template();
      if (path) { toast(`模板已保存: ${path}`, 'success'); }
      return;
    } catch (e) { /* 走兜底 */ }
  }
  // 浏览器模式兜底：用 <a download> 触发下载
  const a = document.createElement('a');
  a.href = '/api/download_template';
  a.download = '发送映射模板.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// ─── 消息开关 ──────────────────────────────────────────────
function toggleMessage() {
  document.getElementById('msgGroup').style.display =
    document.getElementById('enableMsg').checked ? 'block' : 'none';
}

// ─── 预览匹配 ──────────────────────────────────────────────
let previewData = null;      // 缓存最近一次预览结果
let previewView = 'contact'; // 'contact' = 按联系人(发送计划) | 'keyword' = 按关键词(明细)

function escapeFileList(files) {
  return files.length > 0
    ? '<div class="file-list-inline">' + files.map(f => `<div>📄 ${escapeHtml(f)}</div>`).join('') + '</div>'
    : '<span style="color:var(--text-muted);">—</span>';
}

function switchPreviewView(view) {
  previewView = view;
  document.getElementById('vsContact').classList.toggle('active', view === 'contact');
  document.getElementById('vsKeyword').classList.toggle('active', view === 'keyword');
  renderPreview();
}

function renderPreview() {
  if (!previewData) return;
  const data = previewData;
  const tbody = document.getElementById('previewBody');
  const empty = document.getElementById('previewEmpty');
  const head = document.getElementById('previewHead');
  const rows = previewView === 'contact' ? data.grouped : data.matched;

  if (!rows || rows.length === 0) {
    tbody.innerHTML = '';
    empty.style.display = 'block';
    head.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  if (previewView === 'contact') {
    head.innerHTML = '<tr><th style="width:18%">联系人</th><th style="width:22%">命中关键词</th><th>匹配文件（已去重）</th><th style="width:70px;">状态</th></tr>';
    tbody.innerHTML = rows.map(g => {
      const kwHtml = g.keywords.length > 0
        ? '<div class="kw-tags">' + g.keywords.map(k => `<span class="kw-tag">${escapeHtml(k)}</span>`).join('') + '</div>'
        : '<span style="color:var(--text-muted);">—</span>';
      const badge = g.files.length > 0
        ? '<span class="badge badge-info">待发送</span>'
        : '<span class="badge badge-warn">无匹配</span>';
      return `<tr><td>${escapeHtml(g.name)}</td><td>${kwHtml}</td><td>${escapeFileList(g.files)}</td><td>${badge}</td></tr>`;
    }).join('');
  } else {
    head.innerHTML = '<tr><th>关键词</th><th style="width:18%">联系人</th><th>匹配文件</th><th style="width:70px;">状态</th></tr>';
    tbody.innerHTML = rows.map(m => {
      const badge = m.files.length > 0
        ? '<span class="badge badge-info">待发送</span>'
        : '<span class="badge badge-warn">无匹配</span>';
      return `<tr><td>${escapeHtml(m.keyword)}</td><td>${escapeHtml(m.name)}</td><td>${escapeFileList(m.files)}</td><td>${badge}</td></tr>`;
    }).join('');
  }
}

async function previewMatches() {
  const path = document.getElementById('folderPath').value.trim();
  if (!path) return toast('请先输入文件夹路径', 'error');

  const res = await fetch('/api/preview', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({path})
  });
  const data = await res.json();

  if (data.error) return toast(data.error, 'error');

  previewData = data;
  document.getElementById('previewCard').style.display = 'block';
  renderPreview();
  renderAllocStats(data);

  // 合并提示
  const tip = document.getElementById('mergeTip');
  if (data.dup_keywords > 0 || data.dup_files > 0) {
    tip.innerHTML = `已自动合并：<b>${data.dup_keywords}</b> 条重复映射、<b>${data.dup_files}</b> 个重复文件（同一联系人只发一次）`;
  } else {
    tip.textContent = '';
  }

  document.getElementById('btnSend').disabled = data.sendable === 0;
  const statusEl = document.getElementById('sendStatus');
  if (data.unmatched_count > 0) {
    statusEl.innerHTML = `文件 <b style="color:var(--danger)">${data.assigned_count}/${data.folder_file_total}</b> 已分配`
      + ` · ⚠ ${data.unmatched_count} 个未匹配`;
    statusEl.style.color = 'var(--danger)';
  } else {
    statusEl.innerHTML = `映射 ${data.matched_total} 条 → ${data.sendable} 个联系人`
      + ` · 文件 <b>${data.assigned_count}/${data.folder_file_total}</b> 全部已分配`;
    statusEl.style.color = '';
  }
}

// ─── 文件分配统计 ──────────────────────────────────────────
function renderAllocStats(d) {
  const total = d.folder_file_total ?? 0;
  const assigned = d.assigned_count ?? 0;
  const unmatched = d.unmatched_count ?? 0;
  const contacts = d.sendable ?? 0;
  const cov = d.coverage ?? 0;

  document.getElementById('allocTotal').textContent = total;
  document.getElementById('allocAssigned').textContent = assigned;
  document.getElementById('allocContacts').textContent = contacts;
  document.getElementById('allocCoverage').textContent = cov + '%';

  // 未匹配计数
  const uw = document.getElementById('allocUnmatchedWrap');
  const box = document.getElementById('unmatchedBox');
  if (unmatched > 0) {
    uw.style.display = 'block';
    document.getElementById('allocUnmatched').textContent = unmatched;
    box.style.display = 'block';
    const list = document.getElementById('unmatchedList');
    const files = d.unmatched_files || [];
    // 默认展开：漏掉的文件是最需要被看见的，不该再藏一次。
    // 清单可能很长，CSS 里用 max-height + 滚动兜住，不会把页面撑爆。
    list.innerHTML = files.map(f =>
      `<span class="unmatched-item">${escapeHtml(f)}</span>`).join('')
      + (files.length > 20
          ? `<span class="unmatched-more">（共 ${files.length} 个，可滚动查看）</span>`
          : '');
    list.style.display = 'flex';
    document.getElementById('umToggle').textContent = '收起 ▴';
  } else {
    uw.style.display = 'none';
    box.style.display = 'none';
  }

  // 覆盖率进度条
  const fill = document.getElementById('allocCovFill');
  fill.style.width = Math.min(100, cov) + '%';
  fill.className = 'cov-fill' + (cov >= 100 ? '' : (cov >= 80 ? ' warn' : ' danger'));
}

function toggleUnmatched() {
  const list = document.getElementById('unmatchedList');
  const t = document.getElementById('umToggle');
  const shown = list.style.display !== 'none';
  list.style.display = shown ? 'none' : 'flex';
  t.textContent = shown ? '展开 ▾' : '收起 ▴';
}

// ─── 发送 ──────────────────────────────────────────────────
async function startSend() {
  const path = document.getElementById('folderPath').value.trim();
  const enableMsg = document.getElementById('enableMsg').checked;
  const msg = document.getElementById('msgText').value;
  const maxRetries = parseInt(document.getElementById('maxRetries').value, 10) || 2;

  const res = await fetch('/api/send', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({path, enable_msg: enableMsg, msg, max_retries: maxRetries})
  });
  const data = await res.json();
  if (data.error) return toast(data.error, 'error');

  document.getElementById('btnSend').disabled = true;
  document.getElementById('btnStop').disabled = false;
  document.getElementById('btnPreview').disabled = true;
  document.getElementById('progressArea').style.display = 'block';
  document.getElementById('statsCard').style.display = 'none';
  document.getElementById('reviewCard').style.display = 'none';

  toast('已开始发送', 'success');
  pollStatus();
}

async function stopSend() {
  await fetch('/api/stop', { method: 'POST' });
  document.getElementById('btnStop').disabled = true;
  toast('正在停止...', 'error');
}

// ─── 状态轮询 ──────────────────────────────────────────────
let pollTimer = null;
async function pollStatus() {
  const res = await fetch('/api/status');
  const data = await res.json();

  // 日志
  const logEl = document.getElementById('logArea');
  logEl.innerHTML = data.logs.slice(-50).map(l => {
    let cls = '';
    if (l.msg.startsWith('✅')) cls = 'success';
    else if (l.msg.startsWith('❌')) cls = 'fail';
    else if (l.msg.startsWith('⚠')) cls = 'warn';
    else if (l.msg.startsWith('❓')) cls = 'review';
    else if (l.msg.startsWith('📦')) cls = 'summary';
    else if (l.msg.startsWith('ℹ') || l.msg.includes('连接') || l.msg.includes('就绪')) cls = 'info';
    return `<div class="log-line ${cls}">[${l.time}] ${escapeHtml(l.msg)}</div>`;
  }).join('');
  logEl.scrollTop = logEl.scrollHeight;

  // 进度
  const p = data.progress;
  const pct = p.total > 0 ? Math.round(p.current / p.total * 100) : 0;
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressText').textContent = `${p.current} / ${p.total}`;
  document.getElementById('progressPercent').textContent = pct + '%';

  // 待人工确认
  renderReview(data.needs_review || []);

  // 发送汇总
  if (data.stats) renderStats(data.stats);

  // 发送状态
  document.getElementById('sendStatus').textContent =
    data.sending ? `正在发送: ${p.current}/${p.total}` : '';

  if (data.sending) {
    pollTimer = setTimeout(pollStatus, 800);
  } else {
    document.getElementById('btnSend').disabled = false;
    document.getElementById('btnStop').disabled = true;
    document.getElementById('btnPreview').disabled = false;
    if (data.result === 'completed') toast('全部发送完成！', 'success');
    else if (data.result === 'stopped') toast('发送已停止', 'error');
    document.getElementById('sendStatus').textContent = data.result === 'completed' ? '✅ 全部完成' : '';
    // 有待确认条目时提醒用户去处理
    if ((data.needs_review || []).length > 0) {
      toast(`${data.needs_review.length} 个联系人待人工确认`, 'error');
    }
    // 发送结束，把桌面窗口焦点拉回本程序
    if (window.pywebview && window.pywebview.api && window.pywebview.api.bring_to_front) {
      try { window.pywebview.api.bring_to_front(); } catch (e) {}
    }
  }
}

// ─── 待人工确认队列 ─────────────────────────────────────────
function renderReview(items) {
  const card = document.getElementById('reviewCard');
  const list = document.getElementById('reviewList');
  if (!items || items.length === 0) {
    card.style.display = 'none';
    return;
  }
  card.style.display = 'block';
  list.innerHTML = `
    <div class="hint" style="margin:0 0 10px;">
      请到微信里核对以下联系人是否已收到文件，然后选择对应操作。
    </div>` +
    items.map(it => `
      <div class="review-item">
        <div class="review-info">
          <div class="review-name">${escapeHtml(it.name)}</div>
          <div class="review-files">${escapeHtml((it.files || []).join('、'))}</div>
          <div class="review-detail">失败于：${escapeHtml(it.last_result ? it.last_result.stage : '未知')}
            — ${escapeHtml(it.last_result ? it.last_result.detail : '')}</div>
        </div>
        <div class="review-actions">
          <button class="btn btn-sm btn-primary" data-name="${escapeAttr(it.name)}" data-action="confirm"
            onclick="handleReview(this.dataset.name, this.dataset.action)">已发送</button>
          <button class="btn btn-sm btn-outline" data-name="${escapeAttr(it.name)}" data-action="resend"
            onclick="handleReview(this.dataset.name, this.dataset.action)">未发送，重发</button>
        </div>
      </div>`).join('');
}

async function handleReview(name, action) {
  const res = await fetch('/api/retry_confirm', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name, action})
  });
  const data = await res.json();
  if (data.error) return toast(data.error, 'error');
  renderReview(data.needs_review || []);
  if (data.stats) renderStats(data.stats);
  toast(action === 'confirm' ? `已标记「${name}」为发送成功` : `已将「${name}」重新入队`, 'success');
  if (data.hint) toast(data.hint, 'info');
}

// ─── 发送汇总 ──────────────────────────────────────────────
function renderStats(s) {
  const card = document.getElementById('statsCard');
  card.style.display = 'block';

  // 文件维度是用户最关心的——「文件都发出去了吗」
  const hasFileStats = typeof s.total_files === 'number';
  const items = hasFileStats ? [
    { label: '文件总数', value: s.total_files, cls: 'stat-total' },
    { label: '文件成功', value: s.files_done, cls: 'stat-ok' },
    { label: '文件失败', value: s.files_failed, cls: 'stat-fail' },
    { label: '文件待确认', value: s.files_uncertain, cls: 'stat-review' },
    { label: '收件人成功', value: s.succeeded, cls: 'stat-ok' },
    { label: '收件人失败', value: s.failed, cls: 'stat-fail' },
  ] : [
    { label: '总计', value: s.total, cls: 'stat-total' },
    { label: '成功', value: s.succeeded, cls: 'stat-ok' },
    { label: '重试后成功', value: s.retried_then_ok, cls: 'stat-retry' },
    { label: '待人工确认', value: s.needs_review, cls: 'stat-review' },
    { label: '失败', value: s.failed, cls: 'stat-fail' },
  ];
  document.getElementById('statsBody').innerHTML = items
    .map(i => `<div class="stat-box ${i.cls}">
        <div class="stat-value">${i.value}</div>
        <div class="stat-label">${i.label}</div>
      </div>`).join('');

  // 明确的结论：是不是全部文件都成功了
  if (hasFileStats) {
    const ok = s.files_done === s.total_files && s.total_files > 0;
    const uncertain = s.files_uncertain > 0 || s.files_failed > 0;
    const el = document.getElementById('statsConclusion');
    if (ok) {
      el.textContent = `✅ 全部 ${s.total_files} 个文件均已发送成功`;
      el.className = 'stats-conclusion ok';
    } else if (uncertain) {
      el.textContent = `⚠ 有文件未成功发送：失败 ${s.files_failed} 个、待确认 ${s.files_uncertain} 个`
        + `（成功 ${s.files_done}/${s.total_files}）`;
      el.className = 'stats-conclusion warn';
    } else {
      el.textContent = `已发送 ${s.files_done}/${s.total_files} 个文件`;
      el.className = 'stats-conclusion';
    }
    el.style.display = 'block';
  }
}

// ─── 工具函数 ──────────────────────────────────────────────
function toast(msg, type) {
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2500);
}
function escapeHtml(s) {
  const d = document.createElement('div');
  d.textContent = s == null ? '' : String(s);
  return d.innerHTML;
}
// 用于 HTML 属性值（data-name），额外转义引号与尖括号
function escapeAttr(s) {
  return escapeHtml(s).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ─── 拖拽上传 ──────────────────────────────────────────────
const uploadZone = document.getElementById('uploadZone');
uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
  e.preventDefault();
  uploadZone.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) {
    const dt = new DataTransfer();
    dt.items.add(file);
    document.getElementById('excelFile').files = dt.files;
    handleExcelUpload({target: {files: dt.files}});
  }
});

// ─── 启动：恢复上次配置（缺口4）───────────────────────────
loadSettings();
</script>
</body>
</html>'''


# ══════════════════════════════════════════════════════════════
#  API 路由
# ══════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return HTML


@app.route('/api/select_folder', methods=['POST'])
def select_folder():
    """调用 pywebview 原生文件夹选择对话框（桌面模式下由前端直接走 JS API，这里为兜底）"""
    try:
        import webview
        if webview.windows:
            result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
            if result:
                folder = os.path.normpath(result[0])
                _add_log(f"ℹ 已选择文件夹: {folder}")
                return jsonify({"path": folder})
    except Exception as e:
        _add_log(f"⚠ 文件夹选择失败: {e}")
    return jsonify({"path": ""})


@app.route('/api/set_mappings', methods=['POST'])
def set_mappings():
    """手动设置映射表（不走 Excel 上传）"""
    data = request.get_json()
    mappings_raw = data.get('mappings', [])
    mappings = [(item[0].strip(), item[1].strip()) for item in mappings_raw if len(item) == 2 and item[0].strip() and item[1].strip()]
    if not mappings:
        return jsonify({"error": "映射数据为空"})
    state["mappings"] = mappings
    state["excel_uploaded"] = ("手动输入", None)
    _add_log(f"ℹ 手动输入: {len(mappings)} 条映射")
    config_store.flush()          # 映射是最重要的用户资产，立即落盘
    return jsonify({"count": len(mappings)})


@app.route('/api/scan_folder', methods=['POST'])
def scan_folder():
    data = request.get_json()
    folder = data.get('path', '').strip()
    if not folder or not os.path.isdir(folder):
        return jsonify({"error": "文件夹不存在，请检查路径"})

    files = []
    try:
        for name in os.listdir(folder):
            full = os.path.join(folder, name)
            if os.path.isfile(full):
                files.append(name)
    except PermissionError:
        return jsonify({"error": "没有权限访问该文件夹"})

    state["folder_path"] = folder
    config_store.remember_folder(state, folder)
    _add_log(f"ℹ 已选择文件夹: {folder} ({len(files)} 个文件)")
    config_store.flush()          # 选定文件夹是关键时机，立即落盘
    return jsonify({"path": folder, "count": len(files), "files": sorted(files),
                    "recent": state.get("recent_folders", [])})


@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({"error": "请选择文件"})
    file = request.files['file']
    if not file.filename:
        return jsonify({"error": "未选择文件"})

    # secure_filename 会丢掉中文，用时间戳+扩展名保存
    ext = os.path.splitext(file.filename)[1]
    save_name = datetime.now().strftime("%Y%m%d_%H%M%S") + ext
    save_path = os.path.join(UPLOAD_FOLDER, save_name)
    file.save(save_path)

    # 解析 Excel
    try:
        mappings = _read_excel(save_path)
    except Exception as e:
        try: os.remove(save_path)
        except: pass
        return jsonify({"error": f"解析 Excel 失败: {e}"})

    state["excel_uploaded"] = (file.filename, save_path)
    state["mappings"] = mappings
    _add_log(f"ℹ 已上传映射表: {file.filename} ({len(mappings)} 条映射)")
    return jsonify({"filename": file.filename, "count": len(mappings)})


@app.route('/api/download_template')
def download_template():
    """提供映射表模板下载；优先读取编译进 exe 的资源，读取失败则用 openpyxl 动态生成。"""
    src = _resolve_template_source()
    if src:
        return send_file(src, as_attachment=True, download_name=TEMPLATE_NAME,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    buf = _build_template_bytes()
    return send_file(buf, as_attachment=True, download_name=TEMPLATE_NAME,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route('/api/preview', methods=['POST'])
def preview():
    data = request.get_json()
    folder = data.get('path', state.get("folder_path", "")).strip()

    if not folder or not os.path.isdir(folder):
        return jsonify({"error": "请先选择有效的文件夹路径"})
    if not state["mappings"]:
        return jsonify({"error": "请先上传 Excel 映射表"})

    # 扫描文件
    files_in_folder = {}
    for name in os.listdir(folder):
        full = os.path.join(folder, name)
        if os.path.isfile(full):
            files_in_folder[name] = full

    # 逐条匹配（明细，便于核对每个关键词命中了哪些文件）
    matched = []
    for keyword, name in state["mappings"]:
        kw_lower = keyword.lower()
        file_matches = [fn for fn, fp in files_in_folder.items() if kw_lower in fn.lower()]
        matched.append({
            "keyword": keyword,
            "name": name,
            "files": file_matches,
        })

    # 按联系人聚合（真正的发送计划）：
    #   - 多个关键词映射到同一个人 → 合并为一次发送，文件去重，绝不重复发
    #   - 映射表里完全重复的行 → 同样合并
    grouped = []
    grouped_index = {}          # name -> 聚合项
    dup_files = 0
    dup_keywords = 0
    for keyword, name in state["mappings"]:
        kw_lower = keyword.lower()
        file_matches = [fn for fn, fp in files_in_folder.items() if kw_lower in fn.lower()]

        g = grouped_index.get(name)
        if g is None:
            g = {"name": name, "keywords": [], "files": [], "_seen": set()}
            grouped_index[name] = g
            grouped.append(g)
        else:
            dup_keywords += 1

        if keyword not in g["keywords"]:
            g["keywords"].append(keyword)

        for fn in file_matches:
            if fn in g["_seen"]:
                dup_files += 1
                continue
            g["_seen"].add(fn)
            g["files"].append(fn)

    for g in grouped:
        g.pop("_seen", None)    # set 不能序列化

    # ── 分配统计 ──────────────────────────────────────────
    # 用户最关心的问题是「文件夹里的文件是不是都找到人了」。
    # 注意：total_files（已匹配文件数）≠ 文件夹里的文件总数，
    # 没命中任何关键词的文件会被**静默漏掉**——所以必须单独算出来。
    all_file_names = list(files_in_folder.keys())
    folder_file_total = len(all_file_names)

    # 已分配的文件 = 出现在任一联系人的 files 里（合并去重后的实际发送清单）
    assigned = set()
    for g in grouped:
        assigned.update(g["files"])
    assigned_count = len(assigned)

    # 未分配 = 文件夹里有、但没有任何联系人认领
    unmatched = sorted(fn for fn in all_file_names if fn not in assigned)
    unmatched_count = len(unmatched)

    # 覆盖率：已分配文件 / 文件夹总文件
    coverage = (assigned_count / folder_file_total * 100) if folder_file_total else 0.0

    # 统计以「发送计划」为准（所见即所发）
    sendable = sum(1 for g in grouped if g["files"])
    total_files = assigned_count

    state["folder_path"] = folder
    state["matched"] = matched
    state["grouped"] = grouped
    state["unmatched_files"] = unmatched

    log = (f"ℹ 预览完成: {len(matched)} 条映射 → {sendable} 个联系人 | "
           f"文件夹 {folder_file_total} 个文件，已分配 {assigned_count} 个")
    if unmatched_count:
        log += f"，⚠ {unmatched_count} 个文件未匹配到任何人"
    if dup_keywords or dup_files:
        log += f"（已合并 {dup_keywords} 条重复映射、去掉 {dup_files} 个重复文件）"
    _add_log(log)

    return jsonify({
        "matched": matched,
        "matched_total": len(matched),
        "grouped": grouped,
        "grouped_total": len(grouped),
        "dup_keywords": dup_keywords,
        "dup_files": dup_files,
        "sendable": sendable,
        "total_files": total_files,
        # ── 分配统计（新增）──
        "folder_file_total": folder_file_total,
        "assigned_count": assigned_count,
        "unmatched_count": unmatched_count,
        "unmatched_files": unmatched,
        "coverage": round(coverage, 1),
    })


@app.route('/api/send', methods=['POST'])
def send():
    if state["sending"]:
        return jsonify({"error": "正在发送中，请等待完成或先停止"})
    if not state["matched"] and not state["grouped"]:
        return jsonify({"error": "请先预览匹配结果"})

    data = request.get_json()
    enable_msg = data.get("enable_msg", False)
    msg = data.get("msg", "").strip()
    try:
        max_retries = int(data.get("max_retries", 2))
    except (TypeError, ValueError):
        max_retries = 2
    max_retries = max(1, min(5, max_retries))   # 限制在 1~5，避免手滑填 999
    state["max_retries"] = max_retries
    # 附言与重试次数同步进 state，这样它们会被配置持久化，下次启动沿用
    state["enable_msg"] = bool(enable_msg)
    state["msg"] = msg
    config_store.flush()          # 发送前落盘，确保这次用的参数下次还在

    # 发送计划：优先用按联系人聚合去重后的结果（同一人只发一次）
    grouped = state.get("grouped") or []
    if not grouped and state["matched"]:
        # 兜底：从明细现场聚合，保证不重复发送
        idx_map = {}
        for m in state["matched"]:
            g = idx_map.get(m["name"])
            if g is None:
                g = {"name": m["name"], "keywords": [], "files": []}
                idx_map[m["name"]] = g
                grouped.append(g)
            if m["keyword"] not in g["keywords"]:
                g["keywords"].append(m["keyword"])
            for fn in m["files"]:
                if fn not in g["files"]:
                    g["files"].append(fn)
        state["grouped"] = grouped

    send_list = [g for g in grouped if g["files"]]

    if not send_list:
        return jsonify({"error": "没有可发送的文件"})

    state["sending"] = True
    state["stop_requested"] = False
    state["logs"] = []
    state["progress"] = {"current": 0, "total": len(send_list)}
    state["result"] = None
    state["stats"] = None
    state["queue"] = []
    state["needs_review"] = []

    thread = threading.Thread(
        target=_send_worker,
        args=(send_list, enable_msg, msg, max_retries),
        daemon=True
    )
    thread.start()

    return jsonify({"ok": True, "total": len(send_list), "max_retries": max_retries})


@app.route('/api/retry_confirm', methods=['POST'])
def retry_confirm():
    """
    人工处理「状态不明」条目。

    action='confirm' → 用户核对微信后确认已发出 → 标记成功
    action='resend'  → 用户确认没发出去 → 重置并重新入队（需再次点开始发送）
    """
    if state["sending"]:
        return jsonify({"error": "正在发送中，请先等待本轮结束"})

    data = request.get_json() or {}
    name = (data.get("name") or "").strip()
    action = (data.get("action") or "").strip()

    if not name:
        return jsonify({"error": "缺少联系人名称"})
    if action not in ("confirm", "resend"):
        return jsonify({"error": "action 必须是 confirm 或 resend"})

    q = state.get("active_queue")
    # 发送结束后 active_queue 会置空，这里从落盘的 queue 重建一个最小队列
    if q is None:
        q = _rebuild_queue_from_state()

    ok = q.confirm_sent(name) if action == 'confirm' else q.resend(name)
    if not ok:
        return jsonify({"error": f"未找到待确认条目: {name}"})

    state["queue"] = q.to_list()
    state["needs_review"] = [e.to_dict() for e in q.needs_review()]
    state["stats"] = q.stats()
    state["active_queue"] = q

    _add_log(f"ℹ 人工处理「{name}」: {'确认已发送' if action == 'confirm' else '标记为未发送，将重新发送'}")
    remaining = len(q.pending())
    return jsonify({
        "ok": True,
        "needs_review": state["needs_review"],
        "pending_count": remaining,
        "hint": ("队列已全部处理完毕" if remaining == 0
                 else f"有 {remaining} 个待发条目，请重新点「开始发送」"),
    })


def _rebuild_queue_from_state():
    """发送线程结束后，从落盘的 queue 数据重建 RetryQueue，供人工复核操作使用。"""
    q = RetryQueue(policy=RetryPolicy(max_attempts=max(1, state.get("max_retries", 2))))
    for d in state.get("queue", []):
        e = q.add(d["name"], d.get("files", []), d.get("keywords", []), d.get("message"))
        e.status = EntryStatus(d.get("status", "pending"))
        e.attempt = d.get("attempt", 0)
        e.max_attempts = d.get("max_attempts", q.policy.max_attempts)
        e.history = list(d.get("history", []))
        e.last_result = d.get("last_result")
    return q


@app.route('/api/settings', methods=['GET'])
def get_settings():
    """返回可持久化的设置，供前端初始化界面。"""
    return jsonify({
        "folder_path": state.get("folder_path", ""),
        "mappings": state.get("mappings", []),
        "enable_msg": state.get("enable_msg", False),
        "msg": state.get("msg", ""),
        "max_retries": state.get("max_retries", 2),
        "recent_folders": state.get("recent_folders", []),
        "restored": state.get("config_restored", False),
        "warnings": state.get("config_warnings", []),
        "config_path": config_store.path,
        "config_exists": os.path.isfile(config_store.path),
    })


@app.route('/api/settings', methods=['POST'])
def save_settings():
    """保存设置（附言开关/模板/重试次数等）。"""
    data = request.get_json() or {}
    if "enable_msg" in data:
        state["enable_msg"] = bool(data["enable_msg"])
    if "msg" in data:
        state["msg"] = str(data.get("msg") or "")
    if "max_retries" in data:
        try:
            state["max_retries"] = max(1, min(5, int(data["max_retries"])))
        except (TypeError, ValueError):
            pass
    if "window_geometry" in data:
        state["window_geometry"] = str(data["window_geometry"])

    config_store.flush()
    return jsonify({"ok": True, "saved": config_store.info()})


@app.route('/api/settings', methods=['DELETE'])
def clear_settings():
    """清除已保存的配置（恢复出厂）。"""
    if state["sending"]:
        return jsonify({"error": "正在发送中，请先等待完成"})
    # 内存状态回到初始值（保留 logs 便于用户看到刚才发生了什么）
    for k, v in INITIAL_STATE_SNAPSHOT.items():
        if k != "logs":
            state[k] = list(v) if isinstance(v, list) else (
                dict(v) if isinstance(v, dict) else v)
    ok = config_store.delete()
    _add_log("ℹ 已清除保存的配置")
    return jsonify({"ok": ok, "config": config_store.info()})


@app.route('/api/stop', methods=['POST'])
def stop():
    state["stop_requested"] = True
    _add_log("⏹ 用户请求停止...")
    return jsonify({"ok": True})


@app.route('/api/status')
def status():
    return jsonify({
        "sending": state["sending"],
        "logs": state["logs"][-100:],
        "progress": state["progress"],
        "result": state["result"],
        # ── 重试相关 ──
        "needs_review": state["needs_review"],
        "stats": state["stats"],
        "max_retries": state["max_retries"],
    })


# ══════════════════════════════════════════════════════════════
#  核心逻辑
# ══════════════════════════════════════════════════════════════

def _read_excel(filepath):
    """读取 Excel/CSV 映射表"""
    ext = os.path.splitext(filepath)[1].lower()
    mappings = []

    if ext in ('.xlsx', '.xls'):
        wb = load_workbook(filepath, read_only=True, data_only=True)
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True)):
            if i == 0:
                continue
            if not row or len(row) < 2:
                continue
            kw = str(row[0]).strip() if row[0] else ""
            name = str(row[1]).strip() if row[1] else ""
            if kw and name:
                mappings.append((kw, name))
        wb.close()
    elif ext == '.csv':
        import csv
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                if i == 0:
                    continue
                if len(row) < 2:
                    continue
                kw = row[0].strip()
                name = row[1].strip()
                if kw and name:
                    mappings.append((kw, name))
    return mappings


def _send_worker(send_list, enable_msg, msg_template, max_retries=2):
    """后台发送线程 —— 键盘模拟方案，带失败重试与人工复核队列"""
    import pythoncom
    pythoncom.CoInitialize()

    total = len(send_list)

    # ── 构建重试队列 ──
    policy = RetryPolicy(max_attempts=max(1, int(max_retries)))
    queue = RetryQueue(policy=policy)

    def _msg_for(name):
        if enable_msg and msg_template:
            return msg_template.replace("{name}", name).strip()
        return None

    queue.extend(send_list, message_for=_msg_for)
    state["active_queue"] = queue
    state["queue"] = queue.to_list()
    state["needs_review"] = []

    # ── 获取文件夹内文件的全路径 ──
    folder = state["folder_path"]
    file_paths_cache = {}
    if folder and os.path.isdir(folder):
        for name in os.listdir(folder):
            fp = os.path.join(folder, name)
            if os.path.isfile(fp):
                file_paths_cache[name] = fp

    # ── 连接微信 ──
    _add_log("正在连接微信客户端（键盘模拟模式）...")
    sender = WeChatSender()
    sender_set_log(_add_log)

    if not sender.connect():
        _add_log("⚠ 请确保微信PC版已登录且窗口未被最小化")
        state["sending"] = False
        state["result"] = "stopped"
        state["stats"] = None
        pythoncom.CoUninitialize()
        return

    _add_log("✅ 微信连接成功，开始发送...")
    if policy.max_attempts > 1:
        _add_log(f"ℹ 失败自动重试已开启，最多尝试 {policy.max_attempts} 次")

    success = 0
    fail = 0
    stopped = False

    # ── 文件维度计数 ──────────────────────────────────────
    # 用户问的是「文件都发出去了吗」，而联系人维度回答不了这个问题
    # （一个联系人可能带 3 个文件）。所以额外维护文件级进度。
    total_files = sum(len(e.files) for e in queue.entries)
    files_done = 0            # 已成功发送的文件数
    files_failed = 0          # 明确失败的文件数
    files_uncertain = 0       # 状态不明的文件数（需人工确认）
    # 游标：已处理到的文件位置（含失败）。不能拿 files_done 推算——
    # 失败的文件不推进 files_done，会让下一个条目的文件区间重复。
    file_cursor = 0

    # 逐个处理；因重试会原地重跑，故用索引而非直接迭代
    idx = 0
    while idx < len(queue.entries):
        if state["stop_requested"]:
            stopped = True
            _add_log(f"⏹ 已停止发送 (已处理 {idx}/{total} 个联系人, "
                     f"文件 {files_done}/{total_files})")
            state["result"] = "stopped"
            break

        entry = queue.entries[idx]
        idx += 1

        # 已成功/已判定的条目直接跳过（重试在同一次 attempt 内完成）
        if entry.status in (EntryStatus.SUCCEEDED, EntryStatus.FAILED,
                            EntryStatus.NEEDS_REVIEW):
            continue

        n_files = len(entry.files)
        # 文件区间用「游标」算：游标含失败的文件，保证区间不重复、不跳号。
        # 单个文件用 "3/4"，多个用 "1-2/4"——单个用区间读起来别扭
        if n_files == 1:
            file_range = f"文件 {file_cursor + 1}/{total_files}"
        elif n_files > 1:
            file_range = f"文件 {file_cursor + 1}-{file_cursor + n_files}/{total_files}"
        else:
            file_range = f"文件 {file_cursor}/{total_files}"
        file_cursor += n_files

        # 组装文件全路径
        missing = [fn for fn in entry.files if fn not in file_paths_cache]
        if missing:
            r = SendResult.permanent_fail(Stage.FILES, f'文件未找到: {missing[0]}')
            action = queue.handle(entry, r)
        else:
            if len(entry.files) == 1:
                file_paths = file_paths_cache[entry.files[0]]
            else:
                file_paths = [file_paths_cache[fn] for fn in entry.files]

            kw_tag = f"[{'、'.join(entry.keywords)}] " if entry.keywords else ""
            head = (f"  [{idx}/{total}] {entry.name} ← {kw_tag}{entry.display_files}"
                    f"　({file_range})")
            _add_log(f"{head}")

            # ── 重试循环：同一联系人连续尝试，直到成功/不可重试/需复核 ──
            while True:
                if entry.attempt:
                    _add_log(f"  ↻ 第 {entry.attempt + 1} 次尝试 "
                             f"(最多 {entry.max_attempts} 次)")
                try:
                    r = sender.send(entry.name, message=entry.message,
                                    file_paths=file_paths)
                except Exception as e:
                    r = SendResult.safe_retry(Stage.FILES, f'未预期异常: {e}')

                action = queue.handle(entry, r, sleep_fn=None)

                if action != 'retry':
                    break
                # 可安全重试 → 退避后原地再试
                delay = policy.delay_for(entry)
                _add_log(f"  ⚠ {r.stage.value}失败: {r.detail}"
                         f"，{delay:.0f}s 后重试"
                         f" (剩余 {entry.attempts_left} 次)")
                if state["stop_requested"]:
                    break
                time.sleep(delay)

        # ── 按动作输出日志 ──
        if action == 'success':
            files_done += n_files
            if entry.attempt > 1:
                _add_log(f"✅ [{idx}/{total}] {entry.name} ← {entry.display_files} "
                         f"({file_range}，重试 {entry.attempt - 1} 次后成功)")
            else:
                _add_log(f"✅ [{idx}/{total}] {entry.name} ← {entry.display_files} "
                         f"({file_range})")
            success += 1

        elif action == 'review':
            files_uncertain += n_files
            _add_log(f"❓ [{idx}/{total}] {entry.name} 状态不明：{r.detail}。"
                     f"({file_range}) 回车已按下，文件可能已发出，"
                     f"为避免重复发送不自动重试，请人工确认")

        else:  # failed（含重试耗尽）
            files_failed += n_files
            _add_log(f"❌ [{idx}/{total}] {entry.name} 发送失败: {r.detail} "
                     f"({file_range})")
            fail += 1

        state["queue"] = queue.to_list()
        state["needs_review"] = [e.to_dict() for e in queue.needs_review()]
        state["progress"]["current"] = idx

        # 间隔等待
        if idx < len(queue.entries) and not state["stop_requested"]:
            time.sleep(2)

    # ── 汇总 ──
    stats = queue.stats()
    # 补上文件维度的统计（queue.stats() 只有联系人维度）
    stats['total_files'] = total_files
    stats['files_done'] = files_done
    stats['files_failed'] = files_failed
    stats['files_uncertain'] = files_uncertain
    state["stats"] = stats
    state["queue"] = queue.to_list()
    state["needs_review"] = [e.to_dict() for e in queue.needs_review()]

    if state["result"] is None:
        state["result"] = "stopped" if stopped else "completed"

    _add_log("═══════════════════════")
    _add_log(f"📦 文件汇总: 成功 {files_done} / 失败 {files_failed} / "
             f"待确认 {files_uncertain} / 共 {total_files} 个文件")
    if stats['retried_then_ok']:
        _add_log(f"ℹ {stats['retried_then_ok']} 个联系人经重试后成功")
    if stats['needs_review']:
        _add_log(f"❓ {stats['needs_review']} 个状态不明，请在下方「待人工确认」中处理")
    _add_log(f"发送完成: 联系人 成功 {stats['succeeded']} / 失败 {stats['failed']} / 共 {total}")

    # 明确提示是否有文件没发出去——这正是用户关心的
    if files_failed or files_uncertain:
        _add_log(f"⚠ 有文件未成功发送: 失败 {files_failed} 个、待确认 {files_uncertain} 个")
    elif files_done == total_files and total_files > 0:
        _add_log(f"✅ 全部 {total_files} 个文件均已发送成功")

    state["sending"] = False
    # 注意：active_queue 故意保留，供 /api/retry_confirm 做人工复核。
    # 下一次 /api/send 会重新赋值，不会与旧队列混淆。

    pythoncom.CoUninitialize()


# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    import webbrowser
    import threading

    print("=" * 50)
    print("  微信文件群发工具 — Web 版")
    print("  发送模式: 键盘模拟 (适用于所有微信版本)")
    print("  浏览器即将打开: http://localhost:5890")
    print("  关闭此窗口即可退出程序")
    print("=" * 50)

    # 启动后自动打开浏览器
    def _open_browser():
        time.sleep(1.5)
        webbrowser.open('http://localhost:5890')

    threading.Thread(target=_open_browser, daemon=True).start()

    app.run(host='127.0.0.1', port=5890, debug=False)
