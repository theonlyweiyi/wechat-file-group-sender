#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
前端与 API 冒烟测试（缺口2 · 失败重试）
=========================================
用Flask test_client 验证页面与新增 API 可用，并检查前端 JS 里
新增的函数/元素是否都在（防止改HTML 时漏改 JS 导致按钮点了没反应）。

直接运行： python test_api_smoke.py
"""

import os
import re
import sys
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 必须在 import web_app 之前：把配置目录指向临时目录，
# 否则测试会往用户真实的 %APPDATA%\WeChatFileSender 写配置。
import tempfile
os.environ['WFS_CONFIG_DIR'] = tempfile.mkdtemp(prefix='wfs_test_')


def install_stubs():
    if 'pyautogui' in sys.modules:
        return
    pg = types.ModuleType('pyautogui')
    pg.FAILSAFE = False
    pg.PAUSE = 0
    for fn in ('press', 'hotkey', 'write', 'click', 'moveTo', 'keyDown', 'keyUp'):
        setattr(pg, fn, lambda *a, **k: None)
    sys.modules['pyautogui'] = pg

    wc = types.ModuleType('win32clipboard')
    wc.CF_UNICODETEXT = 13
    wc.CF_HDROP = 15
    for fn in ('OpenClipboard', 'CloseClipboard', 'EmptyClipboard', 'SetClipboardData'):
        setattr(wc, fn, lambda *a, **k: None)
    wc.GetClipboardData = lambda *a: None
    sys.modules['win32clipboard'] = wc

    pc = types.ModuleType('pythoncom')
    pc.CoInitialize = lambda: None
    pc.CoUninitialize = lambda: None
    sys.modules['pythoncom'] = pc


install_stubs()

import web_app  # noqa: E402
from web_app import state, HTML  # noqa: E402


class TestPageRenders(unittest.TestCase):

    def setUp(self):
        self.client = web_app.app.test_client()

    def test_index_returns_html(self):
        r = self.client.get('/')
        self.assertEqual(r.status_code, 200)
        self.assertIn('微信文件群发工具', r.get_data(as_text=True))

    def test_max_retries_selector_exists(self):
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn('id="maxRetries"', html, '缺少重试次数选择器')
        # 五个档位
        for v in ('1', '2', '3', '4', '5'):
            self.assertIn(f'value="{v}"', html)

    def test_review_card_exists(self):
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn('id="reviewCard"', html)
        self.assertIn('id="reviewList"', html)

    def test_stats_card_exists(self):
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn('id="statsCard"', html)
        self.assertIn('id="statsBody"', html)

    def test_new_js_functions_defined(self):
        """新增的 JS 函数必须存在，否则按钮点了会报 undefined"""
        html = self.client.get('/').get_data(as_text=True)
        for fn in ('renderReview', 'handleReview', 'renderStats', 'escapeAttr'):
            self.assertIn(f'function {fn}(', html, f'缺少函数 {fn}')

    def test_new_css_classes_defined(self):
        html = self.client.get('/').get_data(as_text=True)
        for cls in ('.review-item', '.stat-box', '.log-line.review'):
            self.assertIn(cls, html, f'缺少样式 {cls}')

    def test_startsend_sends_max_retries(self):
        """startSend 必须把重试次数传给后端"""
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn("getElementById('maxRetries')", html)
        self.assertIn('max_retries: maxRetries', html)

    def test_pollstatus_renders_review_and_stats(self):
        html = self.client.get('/').get_data(as_text=True)
        poll = html.split('async function pollStatus')[1].split('\nasync function')[0]
        self.assertIn('renderReview', poll, 'pollStatus 应调用 renderReview')
        self.assertIn('renderStats', poll, 'pollStatus 应调用 renderStats')

    def test_review_log_marked(self):
        """❓ 开头的日志应有专属样式类"""
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn("startsWith('❓')", html)

    def test_review_buttons_use_dataset(self):
        """按钮应用 dataset 传参，避免引号注入"""
        html = self.client.get('/').get_data(as_text=True)
        self.assertIn('data-name="${escapeAttr(it.name)}"', html)
        self.assertIn('handleReview(this.dataset.name', html)


class TestRetryApi(unittest.TestCase):

    def setUp(self):
        self.client = web_app.app.test_client()
        state.clear()
        state.update({
            "folder_path": "", "excel_uploaded": None, "mappings": [],
            "matched": [], "grouped": [], "sending": False, "stop_requested": False,
            "logs": [], "progress": {"current": 0, "total": 0}, "result": None,
            "max_retries": 2, "queue": [], "needs_review": [], "stats": None,
            "active_queue": None,
        })

    def test_status_shape(self):
        data = self.client.get('/api/status').get_json()
        for f in ('sending', 'logs', 'progress', 'result',
                  'needs_review', 'stats', 'max_retries'):
            self.assertIn(f, data)

    def test_retry_confirm_requires_name(self):
        r = self.client.post('/api/retry_confirm', json={'action': 'confirm'})
        self.assertIn('error', r.get_json())

    def test_retry_confirm_validates_action(self):
        r = self.client.post('/api/retry_confirm',
                             json={'name': '张三', 'action': 'delete_everything'})
        self.assertIn('error', r.get_json())

    def test_retry_confirm_on_empty_queue(self):
        r = self.client.post('/api/retry_confirm',
                             json={'name': '张三', 'action': 'confirm'})
        self.assertIn('error', r.get_json())

    def test_retry_confirm_blocked_while_sending(self):
        state['sending'] = True
        r = self.client.post('/api/retry_confirm',
                             json={'name': '张三', 'action': 'confirm'})
        self.assertIn('error', r.get_json())
        state['sending'] = False

    def test_send_clamps_max_retries(self):
        """max_retries 应被夹在 1~5，防止前端传垃圾值"""
        import threading
        orig = threading.Thread

        def fake_thread(target, args=(), daemon=None):
            class T:
                def start(self_inner): pass
            return T()

        threading.Thread = fake_thread
        try:
            for val, expect in [(0, 1), (1, 1), (3, 3), (99, 5), (-5, 1),
                                ('abc', 2), (None, 2)]:
                # 每次 /api/send 都会置 sending=True，需复位才能再次调用
                state['sending'] = False
                state['stop_requested'] = False
                state['grouped'] = [{'name': '张三', 'files': ['a.pdf'],
                                     'keywords': []}]
                r = self.client.post('/api/send', json={'max_retries': val})
                self.assertEqual(
                    r.get_json()['max_retries'], expect,
                    f'输入 {val!r} 应夹到 {expect}')
        finally:
            threading.Thread = orig
            state['sending'] = False

    def test_send_requires_preview(self):
        r = self.client.post('/api/send', json={})
        self.assertIn('error', r.get_json())

    def test_stop_endpoint(self):
        data = self.client.post('/api/stop').get_json()
        self.assertTrue(data['ok'])
        self.assertTrue(state['stop_requested'])


class TestSettingsApi(unittest.TestCase):
    """缺口4：设置持久化的 API 层"""

    def setUp(self):
        self.client = web_app.app.test_client()

    def test_get_settings_shape(self):
        data = self.client.get('/api/settings').get_json()
        for f in ('folder_path', 'mappings', 'enable_msg', 'msg',
                  'max_retries', 'recent_folders', 'restored',
                  'config_path', 'config_exists'):
            self.assertIn(f, data, f'/api/settings 应返回 {f}')

    def test_post_settings_updates_state(self):
        self.client.post('/api/settings', json={
            'enable_msg': True, 'msg': '你好 {name}', 'max_retries': 4})
        self.assertTrue(state['enable_msg'])
        self.assertEqual(state['msg'], '你好 {name}')
        self.assertEqual(state['max_retries'], 4)

    def test_post_settings_clamps_retries(self):
        self.client.post('/api/settings', json={'max_retries': 999})
        self.assertEqual(state['max_retries'], 5)
        self.client.post('/api/settings', json={'max_retries': -1})
        self.assertEqual(state['max_retries'], 1)

    def test_post_settings_ignores_bad_retries(self):
        before = state['max_retries']
        self.client.post('/api/settings', json={'max_retries': 'abc'})
        self.assertEqual(state['max_retries'], before, '非法值应保持原样')

    def test_post_settings_writes_file(self):
        self.client.post('/api/settings', json={'msg': '持久化测试'})
        import os
        self.assertTrue(os.path.isfile(web_app.config_store.path),
                        'POST 设置应触发落盘')

    def test_delete_settings(self):
        self.client.post('/api/settings', json={'msg': '待清除'})
        data = self.client.delete('/api/settings').get_json()
        self.assertTrue(data['ok'])
        self.assertEqual(state['msg'], '', '清除后应回到初始值')
        self.assertEqual(state['max_retries'], 2)

    def test_delete_blocked_while_sending(self):
        state['sending'] = True
        data = self.client.delete('/api/settings').get_json()
        self.assertIn('error', data)
        state['sending'] = False

    def test_settings_persist_across_requests(self):
        """设置应跨请求保持（同一进程内 state 是共享的）"""
        self.client.post('/api/settings', json={'msg': '跨请求'})
        again = self.client.get('/api/settings').get_json()
        self.assertEqual(again['msg'], '跨请求')

    def test_folder_change_saves_config(self):
        """选定文件夹应写入配置并进最近列表"""
        import tempfile
        d = tempfile.mkdtemp()
        data = self.client.post('/api/scan_folder', json={'path': d}).get_json()
        self.assertEqual(data['path'], d)
        self.assertIn(d, data['recent'])
        self.assertIn(os.path.normpath(d), state['recent_folders'])


class TestSettingsUi(unittest.TestCase):
    """前端新增元素与函数"""

    def setUp(self):
        self.html = self.client_html()

    @staticmethod
    def client_html():
        with web_app.app.test_client() as c:
            return c.get('/').get_data(as_text=True)

    def test_restore_banner_exists(self):
        self.assertIn('id="restoreBanner"', self.html)
        self.assertIn('id="restoreText"', self.html)

    def test_recent_folders_ui_exists(self):
        self.assertIn('id="recentRow"', self.html)
        self.assertIn('id="recentFolders"', self.html)

    def test_config_path_hint_exists(self):
        self.assertIn('id="configPathHint"', self.html)

    def test_new_js_functions_defined(self):
        for fn in ('loadSettings', 'renderRecentFolders', 'applyRecentFolder',
                   'hideRestoreBanner', 'clearConfig', 'persistSettings'):
            self.assertIn(f'function {fn}(', self.html, f'缺少函数 {fn}')

    def test_loadsettings_called_on_startup(self):
        """启动时应自动加载配置"""
        tail = self.html.split('</script>')[0].split('</script>')[-1]
        self.assertIn('loadSettings()', tail[-200:],
                      'script 末尾应调用 loadSettings()')

    def test_settings_triggers_wired(self):
        """附言/重试控件改动应触发 persistSettings"""
        self.assertIn('oninput="persistSettings()"', self.html)
        self.assertIn('onchange="persistSettings()"', self.html)

    def test_restore_uses_escapeattr(self):
        """恢复映射时插入 input value，必须用 escapeAttr"""
        js = self.html.split('<script>')[1].split('</script>')[0]
        restore = js.split('async function loadSettings()')[1].split('\n}')[0]
        self.assertIn('escapeAttr(kw)', restore,
                      '恢复的映射值必须走 escapeAttr 防注入')
        self.assertNotIn('escapeHtml(kw)', restore)

    def test_new_css_classes(self):
        for cls in ('.restore-banner', '.recent-row', '.cfg-path'):
            self.assertIn(cls, self.html, f'缺少样式 {cls}')


class TestJsSanity(unittest.TestCase):
    """对内嵌 JS 做轻量静态检查"""

    def setUp(self):
        self.js = HTML.split('<script>')[1].split('</script>')[0]

    def test_no_unescaped_quote_injs_string(self):
        """renderReview 里的模板串不应再直接内插未转义的名字"""
        self.assertNotIn("handleReview('${escapeHtml", self.js)

    def test_all_called_functions_exist(self):
        """
        检查自定义函数都被定义了。
        只校验本模块定义的、或以特定前缀命名的调用，避开内置方法同名问题
        （如 add/remove/toggle 既是自定义函数也是 DOM 方法）。
        """
        defined = set(re.findall(r'function\s+([a-zA-Z_]\w*)\s*\(', self.js))

        # 明确的自定义函数前缀/名称，这些必须存在
        must_define = {
            'renderReview', 'handleReview', 'renderStats',
            'escapeHtml', 'escapeAttr', 'toast',
            'switchMappingTab', 'addMappingRow', 'removeMappingRow',
            'saveManualMappings', 'selectFolder', 'scanFolder',
            'handleExcelUpload', 'downloadTemplate', 'toggleMessage',
            'switchPreviewView', 'renderPreview', 'previewMatches',
            'startSend', 'stopSend', 'pollStatus',
        }
        missing = must_define - defined
        self.assertFalse(missing, f'缺少自定义函数定义: {missing}')

    def test_no_call_to_undefined_custom_helper(self):
        """renderReview / renderStats / handleReview 内部调用的自定义函数须存在"""
        for fn_name in ('renderReview', 'renderStats', 'handleReview'):
            body = self.js.split(f'function {fn_name}(')[1].split('\n}')[0]
            called = set(re.findall(r'\b([a-zA-Z_]\w*)\s*\(', body))
            defined = set(re.findall(r'function\s+([a-zA-Z_]\w*)\s*\(', self.js))
            custom = {c for c in called
                      if c.startswith(('render', 'escape', 'toast', 'handle'))}
            missing = custom - defined
            self.assertFalse(missing,
                             f'{fn_name} 调用了未定义函数: {missing}')


if __name__ == '__main__':
    unittest.main(verbosity=2)
