#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信文件群发工具 — 桌面版入口
================================
用 pywebview 把 Web 界面渲染进一个原生 Windows 窗口，
不弹浏览器、无控制台。底层使用系统自带的 Edge WebView2。
"""

import os
import sys
import time
import threading
import ctypes
from ctypes import wintypes

import webview
from web_app import app

WINDOW_TITLE = "微信文件群发工具"
DEFAULT_PORT = 5890
PORT_RANGE = 10  # 端口被占用时向后尝试


# ─── Flask 服务线程 ──────────────────────────────────────────
def _run_flask(port):
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


def _free_port():
    import socket
    for p in range(DEFAULT_PORT, DEFAULT_PORT + PORT_RANGE):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', p)) != 0:
                return p
    return DEFAULT_PORT


# ─── pywebview 暴露给前端的 API ───────────────────────────────
class Api:
    """供 JS 调用的原生功能（文件夹选择、窗口聚焦）。"""

    def select_folder(self):
        """弹出系统文件夹选择对话框，返回路径字符串。"""
        try:
            if webview.windows:
                result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
                if result:
                    return os.path.normpath(result[0])
        except Exception:
            pass
        return None

    def bring_to_front(self):
        """将本程序窗口强制置顶（发送结束后把焦点拉回来）。"""
        try:
            hwnd = _find_window_by_title(WINDOW_TITLE)
            if hwnd:
                _force_foreground(hwnd)
        except Exception:
            pass
        return None


# ─── Win32 窗口操作（复用与微信切换相同的方式）─────────────────
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32


def _find_window_by_title(title):
    result = [None]

    def _enum(hwnd, _):
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            if buf.value == title:
                result[0] = hwnd
                return False  # 停止枚举
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(_enum), 0)
    return result[0]


def _force_foreground(hwnd):
    """绕过 Windows 前台锁，将窗口置顶。"""
    # 先恢复（若是最小化）
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        time.sleep(0.3)

    # 用 ALT 键释放前台锁
    user32.keybd_event(0x12, 0, 0, 0)   # ALT down
    user32.keybd_event(0x12, 0, 2, 0)   # ALT up

    # 绑定输入线程，绕过前台限制
    foreground = user32.GetForegroundWindow()
    fg_tid = user32.GetWindowThreadProcessId(foreground, None)
    target_tid = user32.GetWindowThreadProcessId(hwnd, None)
    cur_tid = kernel32.GetCurrentThreadId()
    attached = False
    if fg_tid and fg_tid != cur_tid:
        user32.AttachThreadInput(target_tid, cur_tid, True)
        user32.AttachThreadInput(fg_tid, cur_tid, True)
        attached = True
    try:
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, 0x0001 | 0x0002)  # SWP_NOSIZE|SWP_NOMOVE
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, 0x0001 | 0x0002)
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(fg_tid, cur_tid, False)
            user32.AttachThreadInput(target_tid, cur_tid, False)


# ─── 主流程 ──────────────────────────────────────────────────
def main():
    port = _free_port()
    flask_thread = threading.Thread(target=_run_flask, args=(port,), daemon=True)
    flask_thread.start()

    # 等待 Flask 就绪
    for _ in range(20):
        try:
            import urllib.request
            urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=1)
            break
        except Exception:
            time.sleep(0.3)

    api = Api()
    webview.create_window(
        WINDOW_TITLE,
        f"http://127.0.0.1:{port}/",
        js_api=api,
        width=1040,
        height=780,
        min_size=(920, 640),
    )
    webview.start()


if __name__ == '__main__':
    main()
