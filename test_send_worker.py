#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发送线程集成测试（缺口2 · 失败重试）
=====================================
验证 _send_worker 的真实行为，而不只是 retry.py 的纯逻辑。

本机没有 pyautogui / pywin32 / 微信，所以用 **stub 模块** 顶替：
  - pyautogui：记录按键序列，可按脚本抛异常模拟失败
  - win32clipboard / pythoncom：空实现
  - WeChatSender.connect：直接返回成功
这样不需要真实微信环境就能跑通整条链路。

直接运行： python test_send_worker.py
"""

import os
import sys
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 必须在 import web_app 之前：把配置目录指向临时目录，
# 否则测试会往用户真实的 %APPDATA%\WeChatFileSender 写配置。
import tempfile
os.environ['WFS_CONFIG_DIR'] = tempfile.mkdtemp(prefix='wfs_test_')


# ══════════════════════════════════════════════════════════════
#  stub 微信依赖（必须在 import web_app 之前装好）
# ══════════════════════════════════════════════════════════════

def install_stubs():
    """注入假的 pyautogui / win32clipboard / pythoncom。"""
    if 'pyautogui' in sys.modules:
        return sys.modules['pyautogui']

    pg = types.ModuleType('pyautogui')
    pg.FAILSAFE = True
    pg.PAUSE = 0

    class _Pressed(Exception):
        pass

    def press(key):
        pg.calls.append(('press', key))
        if pg.fail_on_press:
            raise _Pressed(f'模拟按键失败: {key}')

    def hotkey(*keys):
        pg.calls.append(('hotkey', keys))

    def _no_op(*a, **k):
        pass

    pg.calls = []
    pg.fail_on_press = False
    pg.press = press
    pg.hotkey = hotkey
    pg.write = _no_op
    pg.click = _no_op
    pg.FAILSAFE = False
    sys.modules['pyautogui'] = pg

    wc = types.ModuleType('win32clipboard')
    wc.CF_UNICODETEXT = 13
    wc.CF_HDROP = 15
    wc.OpenClipboard = _no_op
    wc.CloseClipboard = _no_op
    wc.EmptyClipboard = _no_op
    wc.SetClipboardData = _no_op
    wc.GetClipboardData = lambda *a: None
    sys.modules['win32clipboard'] = wc

    pc = types.ModuleType('pythoncom')
    pc.CoInitialize = _no_op
    pc.CoUninitialize = _no_op
    sys.modules['pythoncom'] = pc

    return pg


PG = install_stubs()

# 缩短所有 sleep，让测试瞬间完成
import time as _time
_real_sleep = _time.sleep
_time.sleep = lambda s: None

import web_app  # noqa: E402
from web_app import state, _send_worker, _add_log  # noqa: E402
from send_outcome import SendOutcome, SendResult, Stage  # noqa: E402
from retry import EntryStatus  # noqa: E402
import wechat_sender  # noqa: E402


# ══════════════════════════════════════════════════════════════
#  用可控的假 sender 替换真实 WeChatSender
# ══════════════════════════════════════════════════════════════

class ScriptedSender:
    """
    按预设脚本对不同联系人返回不同结果。
    outcomes: {name: [第1次结果, 第2次结果, ...]}，用尽后重复最后一个。
    """

    def __init__(self, outcomes, connect_ok=True):
        self.outcomes = outcomes
        self.connect_ok = connect_ok
        self.calls = []   # [(name, attempt_index)]

    def connect(self):
        return self.connect_ok

    def send(self, name, message=None, file_paths=None):
        seq = self.outcomes.get(name)
        i = len([c for c in self.calls if c[0] == name])
        self.calls.append((name, i))
        if not seq:
            return SendResult.success(Stage.DONE)
        r = seq[min(i, len(seq) - 1)]
        return r


class BaseWorkerTest(unittest.TestCase):

    def setUp(self):
        # 伪造一个包含目标文件的文件夹
        self.folder = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   '_test_files')
        os.makedirs(self.folder, exist_ok=True)
        for fn in ('a.pdf', 'b.pdf', 'c.pdf'):
            with open(os.path.join(self.folder, fn), 'wb') as f:
                f.write(b'x')

        state.clear()
        state.update({
            "folder_path": self.folder, "excel_uploaded": None, "mappings": [],
            "matched": [], "grouped": [], "sending": False, "stop_requested": False,
            "logs": [], "progress": {"current": 0, "total": 0}, "result": None,
            "max_retries": 2, "queue": [], "needs_review": [], "stats": None,
            "active_queue": None,
        })
        self.original_sender = wechat_sender.WeChatSender

    def tearDown(self):
        wechat_sender.WeChatSender = self.original_sender
        _time.sleep = _real_sleep

    def run_worker(self, send_list, max_retries=2, enable_msg=False, msg=''):
        """跑一次 _send_worker，注入 ScriptedSender"""
        sender = self.scripted
        web_app.WeChatSender = lambda: sender
        # wechat_sender.set_log_fn 由 _send_worker 调用，真实实现即可
        _send_worker(send_list, enable_msg, msg, max_retries)
        return sender


# ══════════════════════════════════════════════════════════════

class TestRetryInWorker(BaseWorkerTest):

    def test_safe_failure_is_retried_then_succeeds(self):
        """核心场景：第一次失败(可安全重试) → 自动重试 → 成功"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.safe_retry(Stage.SEARCH, '窗口未响应'),
                     SendResult.success(Stage.FILES)]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': ['工资单']}]

        sender = self.run_worker(send_list, max_retries=2)

        self.assertEqual(len(sender.calls), 2, '应尝试 2 次')
        s = state['stats']
        self.assertEqual(s['succeeded'], 1)
        self.assertEqual(s['retried_then_ok'], 1, '应识别为重试后成功')
        self.assertEqual(s['failed'], 0)
        self.assertEqual(state['result'], 'completed')

    def test_retry_exhausted_marks_failed(self):
        """一直可安全重试但次数用尽 → 彻底失败"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.safe_retry(Stage.SEARCH, '一直失败')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list, max_retries=3)

        self.assertEqual(len(sender.calls), 3, '应尝试满 3 次')
        self.assertEqual(state['stats']['failed'], 1)
        self.assertEqual(state['stats']['succeeded'], 0)

    def test_max_retries_one_means_no_retry(self):
        """max_retries=1 → 不重试"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.safe_retry(Stage.SEARCH, '失败')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list, max_retries=1)

        self.assertEqual(len(sender.calls), 1, 'max_retries=1 不应重试')
        self.assertEqual(state['stats']['failed'], 1)

    def test_uncertain_is_never_auto_retried(self):
        """
        最重要的一条：状态不明时绝不自动重试，避免重复发文件。
        """
        self.scripted = ScriptedSender({
            '张三': [SendResult.uncertain(Stage.FILES, '回车后无响应')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list, max_retries=5)

        self.assertEqual(len(sender.calls), 1, '状态不明时只尝试 1 次')
        self.assertEqual(state['stats']['needs_review'], 1)
        self.assertEqual(state['stats']['succeeded'], 0)
        self.assertEqual(len(state['needs_review']), 1, '应进入待人工确认')

    def test_permanent_fail_not_retried(self):
        """永久失败（文件不存在）→ 只试 1 次"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.permanent_fail(Stage.FILES, '剪贴板被占用')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list, max_retries=5)

        self.assertEqual(len(sender.calls), 1)
        self.assertEqual(state['stats']['failed'], 1)
        self.assertEqual(state['stats']['needs_review'], 0)

    def test_missing_file_short_circuits_without_sending(self):
        """映射里的文件在磁盘上不存在 → 不调用 send，直接永久失败"""
        self.scripted = ScriptedSender({'张三': [SendResult.success()]})
        send_list = [{'name': '张三', 'files': ['不存在.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list, max_retries=3)

        self.assertEqual(len(sender.calls), 0, '文件不存在时不该尝试发送')
        self.assertEqual(state['stats']['failed'], 1)

    def test_mixed_batch(self):
        """混合：直接成功 / 重试后成功 / 待复核 / 永久失败"""
        self.scripted = ScriptedSender({
            'A': [SendResult.success()],
            'B': [SendResult.safe_retry(Stage.SEARCH), SendResult.success()],
            'C': [SendResult.uncertain(Stage.FILES)],
            'D': [SendResult.permanent_fail(Stage.FILES, '不存在')],
        })
        send_list = [
            {'name': 'A', 'files': ['a.pdf'], 'keywords': []},
            {'name': 'B', 'files': ['b.pdf'], 'keywords': []},
            {'name': 'C', 'files': ['c.pdf'], 'keywords': []},
            {'name': 'D', 'files': ['a.pdf'], 'keywords': []},
        ]

        self.run_worker(send_list, max_retries=2)

        s = state['stats']
        self.assertEqual(s['total'], 4)
        self.assertEqual(s['succeeded'], 2)        # A + B
        self.assertEqual(s['retried_then_ok'], 1)  # B
        self.assertEqual(s['needs_review'], 1)     # C
        self.assertEqual(s['failed'], 1)           # D

    def test_message_template_applied_per_contact(self):
        """附言模板应按联系人替换 {name}"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.success()],
            '李四': [SendResult.success()],
        })
        send_list = [
            {'name': '张三', 'files': ['a.pdf'], 'keywords': []},
            {'name': '李四', 'files': ['b.pdf'], 'keywords': []},
        ]
        self.run_worker(send_list, max_retries=1, enable_msg=True,
                        msg='你好 {name}，附件请查收')

        q = state['active_queue']
        msgs = {e.name: e.message for e in q.entries}
        self.assertEqual(msgs['张三'], '你好 张三，附件请查收')
        self.assertEqual(msgs['李四'], '你好 李四，附件请查收')

    def test_connect_failure_aborts_cleanly(self):
        """连不上微信 → 立即中止，不进入循环"""
        self.scripted = ScriptedSender({}, connect_ok=False)
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]

        sender = self.run_worker(send_list)

        self.assertEqual(len(sender.calls), 0)
        self.assertEqual(state['result'], 'stopped')
        self.assertFalse(state['sending'])

    def test_queue_history_captures_attempts(self):
        """每条记录保留完整尝试历史，便于复盘"""
        self.scripted = ScriptedSender({
            '张三': [SendResult.safe_retry(Stage.SEARCH, '第一次'),
                     SendResult.safe_retry(Stage.SEARCH, '第二次'),
                     SendResult.success(Stage.FILES)]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': ['月报']}]

        self.run_worker(send_list, max_retries=3)

        e = state['active_queue'].entries[0]
        self.assertEqual(e.attempt, 3)
        self.assertEqual(len(e.history), 3)
        self.assertEqual(e.history[0]['detail'], '第一次')
        self.assertEqual(e.history[2]['outcome'], 'success')
        self.assertIs(e.status, EntryStatus.SUCCEEDED)

    def test_state_queue_is_json_serializable(self):
        """state['queue'] 要能直接给前端，前端会JSON 化它"""
        import json
        self.scripted = ScriptedSender({
            '张三': [SendResult.uncertain(Stage.FILES, '未知')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': ['月报']}]

        self.run_worker(send_list, max_retries=2)
        json.dumps(state['queue'])
        json.dumps(state['needs_review'])
        json.dumps(state['stats'])

    def test_logs_mention_retry_and_review(self):
        """日志应明确告知用户「正在重试」和「需人工确认」"""
        self.scripted = ScriptedSender({
            'B': [SendResult.safe_retry(Stage.SEARCH, '窗口未响应'),
                  SendResult.success()],
            'C': [SendResult.uncertain(Stage.FILES, '回车后无响应')],
        })
        send_list = [
            {'name': 'B', 'files': ['b.pdf'], 'keywords': []},
            {'name': 'C', 'files': ['c.pdf'], 'keywords': []},
        ]
        self.run_worker(send_list, max_retries=2)

        logs = '\n'.join(l['msg'] for l in state['logs'])
        self.assertIn('重试', logs)
        self.assertIn('状态不明', logs)
        self.assertIn('人工确认', logs)


class TestReviewApiOps(BaseWorkerTest):
    """通过 web_app 的 API 层验证人工复核操作"""

    def _prepare_review(self):
        self.scripted = ScriptedSender({
            '张三': [SendResult.uncertain(Stage.FILES, '回车后无响应')]
        })
        send_list = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]
        self.run_worker(send_list, max_retries=2)
        self.assertEqual(len(state['needs_review']), 1)

    def test_confirm_sent_marks_succeeded(self):
        self._prepare_review()
        with web_app.app.test_client() as c:
            r = c.post('/api/retry_confirm',
                       json={'name': '张三', 'action': 'confirm'})
            data = r.get_json()
        self.assertTrue(data['ok'])
        self.assertEqual(len(data['needs_review']), 0)
        self.assertEqual(state['stats']['succeeded'], 1)

    def test_resend_requeues(self):
        self._prepare_review()
        with web_app.app.test_client() as c:
            r = c.post('/api/retry_confirm',
                       json={'name': '张三', 'action': 'resend'})
            data = r.get_json()
        self.assertTrue(data['ok'])
        self.assertEqual(data['pending_count'], 1, '应重新入队等待重发')
        q = state['active_queue']
        self.assertIs(q.entries[0].status, EntryStatus.PENDING)
        self.assertEqual(q.entries[0].attempt, 0)

    def test_invalid_action_rejected(self):
        self._prepare_review()
        with web_app.app.test_client() as c:
            r = c.post('/api/retry_confirm',
                       json={'name': '张三', 'action': '乱填'})
            data = r.get_json()
        self.assertIn('error', data)

    def test_unknown_name_rejected(self):
        self._prepare_review()
        with web_app.app.test_client() as c:
            r = c.post('/api/retry_confirm',
                       json={'name': '查无此人', 'action': 'confirm'})
            data = r.get_json()
        self.assertIn('error', data)

    def test_rebuild_queue_from_state(self):
        """active_queue 丢失时应能从 state 重建，复核功能不失效"""
        self._prepare_review()
        state['active_queue'] = None
        q = web_app._rebuild_queue_from_state()
        self.assertEqual(len(q.entries), 1)
        self.assertIs(q.entries[0].status, EntryStatus.NEEDS_REVIEW)
        self.assertEqual(q.entries[0].attempt, 1)

    def test_send_api_accepts_max_retries(self):
        """确认 /api/send 正确接收并夹取 max_retries"""
        with web_app.app.test_client() as c:
            # 造一个可发送的 grouped
            state['grouped'] = [{'name': '张三', 'files': ['a.pdf'], 'keywords': []}]
            # 直接拦起线程，避免真发
            import threading
            started = {}
            orig = threading.Thread

            def fake_thread(target, args=(), daemon=None):
                started['args'] = args
                class T:
                    def start(self_): pass
                return T()
            threading.Thread = fake_thread
            try:
                r = c.post('/api/send', json={'max_retries': 99})
                data = r.get_json()
            finally:
                threading.Thread = orig

        self.assertEqual(data['max_retries'], 5, '应被夹到上限 5')
        self.assertEqual(started['args'][3], 5)

    def test_status_includes_retry_fields(self):
        with web_app.app.test_client() as c:
            data = c.get('/api/status').get_json()
        for f in ('needs_review', 'stats', 'max_retries'):
            self.assertIn(f, data, f'/api/status 应包含 {f}')


if __name__ == '__main__':
    unittest.main(verbosity=2)
