#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信发送模块 —— 键盘模拟 + 剪贴板
==================================
不依赖 UIAutomation，通过 win32 窗口操作 + pyautogui 键盘模拟实现。
适用于所有微信版本（包括 Qt 合成窗口渲染的 4.1+）。

原理：
1. 找到微信窗口 → 恢复/激活
2. Ctrl+F 搜索联系人 → 输入名称 → 回车打开聊天
3. 发送消息（可选）
4. 将文件复制到系统剪贴板(CF_HDROP) → Ctrl+V → 回车发送
"""

import os
import time
import ctypes
import struct
import subprocess

# ── 键盘模拟 ──
import pyautogui
pyautogui.FAILSAFE = False
PAUSE_BETWEEN = 0.05

# ── Win32 API ──
from ctypes import wintypes

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
shell32 = ctypes.windll.shell32

WS_EX_APPWINDOW = 0x40000

# ── 剪贴板 ──
import pythoncom
import win32clipboard

CF_HDROP = 15


# ── 四态定义放在 send_outcome.py（无微信依赖，可独立测试）──
# re-export，供 `from wechat_sender import SendResult` 的旧用法继续可用
from send_outcome import SendOutcome, SendResult, Stage  # noqa: E402


def _add_log(msg):
    """日志回调，默认打印到控制台"""
    print(f"[WeChatSender] {msg}")


_log_fn = _add_log


def set_log_fn(fn):
    global _log_fn
    _log_fn = fn


def log(msg):
    _log_fn(msg)


# ══════════════════════════════════════════════════════════════
#  窗口管理
# ══════════════════════════════════════════════════════════════

def find_wechat_window():
    """查找微信主窗口 HWND，返回 (hwnd, title) 或 (None, None)"""
    found = []

    def enum_callback(hwnd, lparam):
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True
        buf = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buf, length + 1)
        title = buf.value

        class_buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buf, 256)
        clsname = class_buf.value

        # 匹配微信窗口：Qt51514QWindowIcon 或 微信/Weixin 标题
        if 'Qt' in clsname and title in ('微信', 'Weixin'):
            ex_style = user32.GetWindowLongW(hwnd, -20)
            r = wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(r))
            w, h = r.right - r.left, r.bottom - r.top
            is_visible = bool(user32.IsWindowVisible(hwnd))
            found.append((hwnd, title, clsname, w, h, is_visible, ex_style))
        return True

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)

    if not found:
        return None, None

    # 优先选可见的、有 APPWINDOW 标记的、更大的
    def score(item):
        hwnd, title, clsname, w, h, vis, ex = item
        return (vis * 100 + bool(ex & WS_EX_APPWINDOW) * 50 + w * h)

    found.sort(key=score, reverse=True)
    best = found[0]
    return best[0], best[1]


def restore_window(hwnd):
    """恢复最小化的窗口（即使不确定是否最小化也强制恢复）"""
    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    time.sleep(0.5)
    # 确认已恢复
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)
        time.sleep(0.5)


def activate_window(hwnd):
    """激活窗口到前台（强制切换，绕过 Windows 前台锁）"""
    restore_window(hwnd)

    # 方法1：keybd_event 模拟 ALT 键，释放前台锁（最可靠的后台抢前台方式）
    user32.keybd_event(0x12, 0, 0, 0)       # ALT down
    user32.keybd_event(0x12, 0, 0x0002, 0)  # ALT up

    # 方法2：AttachThreadInput 把线程输入队列绑定，骗过 Windows 前台锁
    foreground_hwnd = user32.GetForegroundWindow()
    foreground_tid = user32.GetWindowThreadProcessId(foreground_hwnd, None)
    target_tid = user32.GetWindowThreadProcessId(hwnd, None)
    current_tid = kernel32.GetCurrentThreadId()

    attached = False
    if foreground_tid and foreground_tid != current_tid:
        user32.AttachThreadInput(foreground_tid, current_tid, True)
        attached = True
    if target_tid and target_tid != current_tid:
        user32.AttachThreadInput(target_tid, current_tid, True)

    try:
        # HWND_TOPMOST 置顶 + 显示
        user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0,
                           0x0001 | 0x0002 | 0x0040)  # SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW
        time.sleep(0.1)
        # 取消置顶（恢复正常 z-order）
        user32.SetWindowPos(hwnd, -2, 0, 0, 0, 0,
                           0x0001 | 0x0002)  # SWP_NOMOVE|SWP_NOSIZE
        user32.ShowWindow(hwnd, 5)  # SW_SHOW
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        if attached:
            if foreground_tid and foreground_tid != current_tid:
                user32.AttachThreadInput(foreground_tid, current_tid, False)
            if target_tid and target_tid != current_tid:
                user32.AttachThreadInput(target_tid, current_tid, False)

    time.sleep(0.6)


# ══════════════════════════════════════════════════════════════
#  文件剪贴板 (CF_HDROP + DROPFILES)
# ══════════════════════════════════════════════════════════════

def copy_files_to_clipboard(file_paths):
    """
    将文件列表复制到系统剪贴板(CF_HDROP格式)。
    WeChat Ctrl+V 时会识别为文件发送。
    """
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    # 构造双 null 结尾的文件列表字符串
    file_list = '\0'.join(file_paths) + '\0\0'
    file_list_bytes = file_list.encode('utf-16-le')

    # DROPFILES 结构 (20 bytes)
    df = struct.pack('Iiiii', 20, 0, 0, 0, 1)  # fWide=1 (Unicode)
    data = df + file_list_bytes

    # 方案1: 直接用 pywin32 传 bytes（内部自动分配全局内存）
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(CF_HDROP, data)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        log(f"⚠ pywin32 直接写入失败: {e}，尝试 GlobalAlloc 方案...")

    # 方案2: 手动 GlobalAlloc（修复 64 位指针截断问题）
    try:
        # 关键修复：设置正确的 restype/argtypes，防止 64 位指针被截断为 32 位
        kernel32.GlobalAlloc.restype = ctypes.c_void_p
        kernel32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalLock.restype = ctypes.c_void_p
        kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalUnlock.restype = ctypes.c_int
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        kernel32.GlobalFree.restype = ctypes.c_void_p
        kernel32.GlobalFree.argtypes = [ctypes.c_void_p]

        GMEM_MOVEABLE = 0x0002
        hglobal = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not hglobal:
            log("⚠ GlobalAlloc 失败，尝试 PowerShell 方案...")
            return _copy_files_via_powershell(file_paths)

        ptr = kernel32.GlobalLock(hglobal)
        if not ptr:
            log("⚠ GlobalLock 失败，尝试 PowerShell 方案...")
            kernel32.GlobalFree(hglobal)
            return _copy_files_via_powershell(file_paths)

        ctypes.memmove(ptr, data, len(data))
        kernel32.GlobalUnlock(hglobal)

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(CF_HDROP, hglobal)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        log(f"⚠ GlobalAlloc 方案失败: {e}，尝试 PowerShell 方案...")
        return _copy_files_via_powershell(file_paths)


def _copy_files_via_powershell(file_paths):
    """备选：通过 PowerShell 复制文件到剪贴板"""
    try:
        paths = ','.join(f"'{p}'" for p in file_paths)
        cmd = f'powershell -NoProfile -Command "Set-Clipboard -Path {paths}"'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            return True
        else:
            log(f"⚠ PowerShell 剪贴板失败: {result.stderr.strip()}")
            return False
    except Exception as e:
        log(f"⚠ PowerShell 调用失败: {e}")
        return False


# ══════════════════════════════════════════════════════════════
#  核心发送逻辑
# ══════════════════════════════════════════════════════════════

def _paste_text(text):
    """
    通过剪贴板粘贴中文文本（pyautogui.write 不支持非 ASCII 字符）。
    复制文本到剪贴板 → Ctrl+V 粘贴。
    """
    import win32clipboard
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()
    except Exception:
        # 备选：用 PowerShell 设置剪贴板文本
        subprocess.run(
            ['powershell', '-NoProfile', '-Command', f'Set-Clipboard -Value "{text}"'],
            capture_output=True, timeout=5
        )
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)


class WeChatSender:
    """微信发送器 —— 键盘模拟方案"""

    def __init__(self):
        self.hwnd = None
        self.title = None
        self._connected = False
        # 身份校验钩子：callable(name) -> (bool, str)。缺口2阶段先落为None，
        # 即不做校验；后续「联系人精确匹配」会注入实现，见 verify_current_chat。
        self._verify_hook = None

    def set_verify_hook(self, fn):
        """注入身份校验函数。fn(name) -> (matched: bool, detail: str)"""
        self._verify_hook = fn

    def verify_current_chat(self, name):
        """
        打开会话后校验「当前打开的是不是这个人」。

        当前微信4.1+ 为 Qt Quick 像素渲染，不暴露 UIA 控件树，暂无可用实现。
        未注入钩子时返回 (None, '未启用校验') —— None 表示「无法判断」，
        调用方应区别于「校验不通过」，不要把无法判断当成安全。
        """
        if self._verify_hook is None:
            return None, '未启用校验'
        try:
            return self._verify_hook(name)
        except Exception as e:
            return False, f'校验异常: {e}'

    def connect(self):
        """连接微信窗口"""
        log("正在查找微信窗口...")
        self.hwnd, self.title = find_wechat_window()
        if self.hwnd is None:
            log("❌ 未找到微信窗口，请确保微信PC版已登录")
            return False

        restore_window(self.hwnd)
        activate_window(self.hwnd)

        # 确认窗口已恢复
        r = wintypes.RECT()
        user32.GetWindowRect(self.hwnd, ctypes.byref(r))
        w, h = r.right - r.left, r.bottom - r.top
        log(f"✅ 微信窗口已连接 (hwnd={self.hwnd}, 大小={w}x{h})")
        self._connected = True
        return True

    def _search_and_open_chat(self, name):
        """搜索并打开与指定联系人的聊天"""
        activate_window(self.hwnd)
        time.sleep(0.3)

        # Ctrl+F 打开搜索
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(0.6)

        # 清空搜索框并粘贴联系人名称（pyautogui.write 不支持中文）
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        _paste_text(name)
        time.sleep(1.0)

        # 回车选中第一个结果
        pyautogui.press('enter')
        time.sleep(1.5)

    def send_message(self, text):
        """在当前聊天窗口发送文字消息。返回 SendResult。"""
        if not text:
            return SendResult.success(Stage.MESSAGE, '空附言，跳过')
        try:
            activate_window(self.hwnd)
            time.sleep(0.3)
        except Exception as e:
            return SendResult.safe_retry(Stage.MESSAGE, f'窗口激活失败: {e}')

        _paste_text(text)
        time.sleep(0.3)
        try:
            pyautogui.press('enter')
        except Exception as e:
            # 附言粘贴成功但回车失败 —— 文字还在输入框里，未发出，可安全重试
            return SendResult.safe_retry(Stage.MESSAGE, f'回车失败: {e}')
        time.sleep(0.5)
        return SendResult.success(Stage.MESSAGE)

    def send_files(self, file_paths):
        """在当前聊天窗口发送文件。返回 SendResult。"""
        if isinstance(file_paths, str):
            file_paths = [file_paths]

        # 前置校验：文件必须真实存在，否则重试再多次也没用
        missing = [p for p in file_paths if not os.path.isfile(p)]
        if missing:
            return SendResult.permanent_fail(
                Stage.FILES, f'文件不存在: {os.path.basename(missing[0])}')

        try:
            activate_window(self.hwnd)
            time.sleep(0.3)
        except Exception as e:
            return SendResult.safe_retry(Stage.FILES, f'窗口激活失败: {e}')

        ok = copy_files_to_clipboard(file_paths)
        if not ok:
            # 剪贴板写入失败发生在粘贴之前，文件确定没发出去。
            # 但多半有其它程序占着剪贴板，连续重试通常也是失败
            return SendResult.permanent_fail(Stage.FILES, '无法写入剪贴板')
        time.sleep(0.5)

        try:
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            # 尚未回车，文件确定没发出去，可安全重试
            return SendResult.safe_retry(Stage.FILES, f'粘贴失败: {e}')
        time.sleep(2.0)  # 等待文件加载/预览

        try:
            pyautogui.press('enter')
        except Exception as e:
            # 已粘贴到输入框但回车失败 → 文件还在输入框，可安全重试
            return SendResult.safe_retry(Stage.FILES, f'回车失败: {e}')
        time.sleep(1.0)

        # 回车已按下，消息很可能已发出
        return SendResult.success(Stage.FILES)

    def send(self, name, message=None, file_paths=None):
        """
        给指定联系人发送消息和/或文件。

        返回 SendResult（结构化四态），不再返回裸 bool。
        调用方应检查 result.outcome，而非 if result:
        """
        if not self._connected:
            try:
                ok = self.connect()
            except Exception as e:
                return SendResult.safe_retry(Stage.CONNECT, f'连接异常: {e}')
            if not ok:
                return SendResult.safe_retry(Stage.CONNECT, '未找到微信窗口')

        # 打开聊天
        try:
            self._search_and_open_chat(name)
        except Exception as e:
            return SendResult.safe_retry(Stage.SEARCH, f'搜索异常: {e}')

        # 身份校验（钩子未启用时返回 None，不阻断流程）
        verified, vdetail = self.verify_current_chat(name)
        if verified is False:
            # 明确不匹配 —— 绝不发送
            return SendResult.permanent_fail(Stage.VERIFY, vdetail)
        if verified is None:
            log(f'⚠ 身份未校验: {vdetail}')

        # 发送附言
        if message:
            r = self.send_message(message)
            if not r.ok:
                return r

        # 发送文件
        if file_paths:
            return self.send_files(file_paths)

        return SendResult.success(Stage.DONE)

    def close(self):
        """清理（目前无需特殊清理）"""
        self._connected = False
        self.hwnd = None


# ══════════════════════════════════════════════════════════════
#  测试入口
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    sender = WeChatSender()
    if sender.connect():
        print("连接成功！")
        # 测试发送文字
        r = sender.send("文件传输助手", message="这是一条自动测试消息")
        print(f"结果: {r.outcome.value} /阶段={r.stage.value} / {r.detail}")
    else:
        print("连接失败！")
