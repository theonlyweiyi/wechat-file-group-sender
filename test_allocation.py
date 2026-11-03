#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分配统计与文件计数测试
======================
覆盖用户提出的两个诉求：
  1. 匹配区能看出「文件夹里的文件是否都分配到人了」
  2. 发送日志能看出「文件是否全部发送成功」

直接运行： python test_allocation.py
"""

import os
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 隔离配置目录，避免污染用户真实 %APPDATA%
os.environ['WFS_CONFIG_DIR'] = tempfile.mkdtemp(prefix='wfs_alloc_test_')

# stub 微信依赖
pg = types.ModuleType('pyautogui')
pg.FAILSAFE = False
for _fn in ('press', 'hotkey', 'write', 'click'):
    setattr(pg, _fn, lambda *a, **k: None)
sys.modules['pyautogui'] = pg

wc = types.ModuleType('win32clipboard')
wc.CF_UNICODETEXT = 13
wc.CF_HDROP = 15
for _fn in ('OpenClipboard', 'CloseClipboard', 'EmptyClipboard', 'SetClipboardData'):
    setattr(wc, _fn, lambda *a, **k: None)
sys.modules['win32clipboard'] = wc

pc = types.ModuleType('pythoncom')
pc.CoInitialize = lambda: None
pc.CoUninitialize = lambda: None
sys.modules['pythoncom'] = pc

import web_app  # noqa: E402
from web_app import state, _send_worker, _add_log  # noqa: E402
from send_outcome import SendResult, Stage  # noqa: E402


def make_folder(files):
    """建临时文件夹并放入指定文件，返回路径"""
    d = tempfile.mkdtemp(prefix='wfs_alloc_files_')
    for fn in files:
        with open(os.path.join(d, fn), 'wb') as f:
            f.write(b'x')
    return d


class TestPreviewAllocation(unittest.TestCase):
    """匹配区的分配统计"""

    def setUp(self):
        self.client = web_app.app.test_client()
        state.clear()
        state.update({
            "folder_path": "", "excel_uploaded": None, "mappings": [],
            "matched": [], "grouped": [], "unmatched_files": [],
            "sending": False, "stop_requested": False,
            "logs": [], "progress": {"current": 0, "total": 0}, "result": None,
            "max_retries": 2, "queue": [], "needs_review": [], "stats": None,
            "active_queue": None, "enable_msg": False, "msg": "",
            "recent_folders": [], "window_geometry": None,
            "config_restored": False, "config_warnings": [],
        })

    def _preview(self, files, mappings):
        d = make_folder(files)
        state['mappings'] = mappings
        r = self.client.post('/api/preview', json={'path': d})
        return r.get_json()

    def test_all_files_matched(self):
        """全部文件都匹配到人时，未匹配数为 0、覆盖率 100%"""
        data = self._preview(
            ['张三_工资单.xlsx', '李四_工资单.xlsx'],
            [('工资单', '张三'), ('工资单', '李四')])
        self.assertEqual(data['folder_file_total'], 2)
        self.assertEqual(data['assigned_count'], 2)
        self.assertEqual(data['unmatched_count'], 0)
        self.assertEqual(data['coverage'], 100.0)
        self.assertEqual(data['unmatched_files'], [])

    def test_some_files_unmatched(self):
        """关键场景：有文件没匹配到人，必须能被看见"""
        data = self._preview(
            ['张三_工资单.xlsx', '李四_工资单.xlsx', '无关文件.txt', '另一个.docx'],
            [('工资单', '张三'), ('工资单', '李四')])
        self.assertEqual(data['folder_file_total'], 4)
        self.assertEqual(data['assigned_count'], 2)
        self.assertEqual(data['unmatched_count'], 2)
        self.assertEqual(sorted(data['unmatched_files']),
                         ['another.txt'.replace('another', 'another'), '无关文件.txt']
                         if False else sorted(['无关文件.txt', '另一个.docx']))
        self.assertEqual(data['coverage'], 50.0)

    def test_no_file_matched(self):
        """一个都没匹配上"""
        data = self._preview(['a.txt', 'b.txt'], [('工资单', '张三')])
        self.assertEqual(data['folder_file_total'], 2)
        self.assertEqual(data['assigned_count'], 0)
        self.assertEqual(data['unmatched_count'], 2)
        self.assertEqual(data['coverage'], 0.0)
        self.assertEqual(data['sendable'], 0)

    def test_dedup_does_not_inflate_assigned_count(self):
        """
        关键场景：两个关键词都命中同一个文件时，
        assigned_count 应按去重后的实际文件数算，不能重复计数。
        """
        data = self._preview(
            ['张三_8月工资.xlsx'],
            [('工资', '张三'), ('8月', '张三')])
        self.assertEqual(data['folder_file_total'], 1)
        self.assertEqual(data['assigned_count'], 1, '同一文件只算一次')
        self.assertEqual(data['unmatched_count'], 0)
        self.assertEqual(data['dup_files'], 1)
        self.assertEqual(data['coverage'], 100.0)

    def test_same_contact_multiple_files(self):
        """一个联系人多个文件：文件数正确、联系人数为 1"""
        data = self._preview(
            ['张三_工资.xlsx', '张三_月报.pdf', '李四_合同.docx'],
            [('工资', '张三'), ('月报', '张三'), ('合同', '李四')])
        self.assertEqual(data['folder_file_total'], 3)
        self.assertEqual(data['assigned_count'], 3)
        self.assertEqual(data['sendable'], 2, '张三合并成一次发送')
        self.assertEqual(data['grouped_total'], 2)

    def test_total_files_equals_assigned(self):
        """total_files 现在等于已分配数（向后兼容的旧字段）"""
        data = self._preview(
            ['a_工资.xlsx', 'b_无关.txt'], [('工资', '张三')])
        self.assertEqual(data['total_files'], data['assigned_count'])
        self.assertEqual(data['total_files'], 1)

    def test_empty_folder(self):
        d = make_folder([])
        state['mappings'] = [('工资', '张三')]
        data = self.client.post('/api/preview', json={'path': d}).get_json()
        self.assertEqual(data['folder_file_total'], 0)
        self.assertEqual(data['coverage'], 0.0)

    def test_unmatched_saved_to_state(self):
        """未匹配清单要存进 state，前端/日志才能用"""
        data = self._preview(['x_工资.xlsx', 'y_杂.txt'], [('工资', '张三')])
        self.assertEqual(state['unmatched_files'], ['y_杂.txt'])
        self.assertEqual(len(data['unmatched_files']), 1)

    def test_log_mentions_unmatched(self):
        """日志里要明确提示有文件没匹配上"""
        self._preview(['x_工资.xlsx', 'y_杂.txt'], [('工资', '张三')])
        logs = '\n'.join(l['msg'] for l in state['logs'])
        self.assertIn('未匹配', logs)

    def test_case_insensitive_matching(self):
        """关键词大小写不敏感也要计入已分配"""
        data = self._preview(['Zhang_Report.PDF'], [('report', '张三')])
        self.assertEqual(data['assigned_count'], 1)
        self.assertEqual(data['unmatched_count'], 0)


class ScriptedSender:
    """按脚本返回不同结果，用于验证文件计数"""

    def __init__(self, outcomes, connect_ok=True):
        self.outcomes = outcomes
        self.connect_ok = connect_ok
        self.calls = []

    def connect(self):
        return self.connect_ok

    def send(self, name, message=None, file_paths=None):
        seq = self.outcomes.get(name)
        i = len([c for c in self.calls if c[0] == name])
        self.calls.append((name, i))
        if not seq:
            return SendResult.success(Stage.FILES)
        return seq[min(i, len(seq) - 1)]


class TestSendFileCounting(unittest.TestCase):
    """发送日志的文件维度计数"""

    def setUp(self):
        self.original_sender = web_app.WeChatSender
        state.clear()
        state.update({
            "folder_path": "", "excel_uploaded": None, "mappings": [],
            "matched": [], "grouped": [], "unmatched_files": [],
            "sending": False, "stop_requested": False,
            "logs": [], "progress": {"current": 0, "total": 0}, "result": None,
            "max_retries": 2, "queue": [], "needs_review": [], "stats": None,
            "active_queue": None, "enable_msg": False, "msg": "",
            "recent_folders": [], "window_geometry": None,
            "config_restored": False, "config_warnings": [],
        })

    def tearDown(self):
        web_app.WeChatSender = self.original_sender

    def _run(self, send_list, outcomes, max_retries=2):
        d = make_folder(['f1.txt', 'f2.txt', 'f3.txt'])
        state['folder_path'] = d
        sender = ScriptedSender(outcomes)
        web_app.WeChatSender = lambda: sender
        _send_worker(send_list, False, '', max_retries)
        return sender

    def _logs(self):
        return '\n'.join(l['msg'] for l in state['logs'])

    def test_all_success_file_count(self):
        """全部成功时，文件成功数 == 文件总数"""
        send_list = [
            {'name': '张三', 'files': ['f1.txt', 'f2.txt'], 'keywords': []},
            {'name': '李四', 'files': ['f3.txt'], 'keywords': []},
        ]
        self._run(send_list, {})
        s = state['stats']
        self.assertEqual(s['total_files'], 3)
        self.assertEqual(s['files_done'], 3)
        self.assertEqual(s['files_failed'], 0)
        self.assertEqual(s['files_uncertain'], 0)

    def test_log_shows_file_progress(self):
        """日志里要能看到文件维度的进度"""
        send_list = [
            {'name': '张三', 'files': ['f1.txt', 'f2.txt'], 'keywords': []},
            {'name': '李四', 'files': ['f3.txt'], 'keywords': []},
        ]
        self._run(send_list, {})
        logs = self._logs()
        self.assertIn('文件 1-2/3', logs)
        self.assertIn('文件 3/3', logs)

    def test_failure_counts_files_not_contacts(self):
        """一个联系人带 3 个文件失败，应记 3 个文件失败"""
        send_list = [
            {'name': '张三', 'files': ['f1.txt', 'f2.txt', 'f3.txt'], 'keywords': []},
        ]
        self._run(send_list, {
            '张三': [SendResult.permanent_fail(Stage.FILES, '剪贴板被占')]})
        s = state['stats']
        self.assertEqual(s['files_failed'], 3, '按文件计数而非联系人')
        self.assertEqual(s['files_done'], 0)
        self.assertEqual(s['failed'], 1, '联系人维度仍是 1')

    def test_uncertain_counts_files(self):
        """状态不明的文件也要计入"""
        send_list = [
            {'name': '张三', 'files': ['f1.txt', 'f2.txt'], 'keywords': []},
        ]
        self._run(send_list, {
            '张三': [SendResult.uncertain(Stage.FILES, '回车后无响应')]})
        s = state['stats']
        self.assertEqual(s['files_uncertain'], 2)
        self.assertEqual(s['files_done'], 0)

    def test_mixed_outcome_file_counts(self):
        """混合场景：成功 2 个文件、失败 1 个、待确认 1 个"""
        send_list = [
            {'name': 'A', 'files': ['f1.txt', 'f2.txt'], 'keywords': []},
            {'name': 'B', 'files': ['f3.txt'], 'keywords': []},
        ]
        self._run(send_list, {
            'A': [SendResult.success(Stage.FILES)],
            'B': [SendResult.permanent_fail(Stage.FILES, '失败')],
        })
        s = state['stats']
        self.assertEqual(s['total_files'], 3)
        self.assertEqual(s['files_done'], 2)
        self.assertEqual(s['files_failed'], 1)

    def test_summary_log_has_file_totals(self):
        """汇总日志里要有文件维度统计"""
        send_list = [{'name': '张三', 'files': ['f1.txt'], 'keywords': []}]
        self._run(send_list, {})
        logs = self._logs()
        self.assertIn('📦', logs)
        self.assertIn('文件汇总', logs)
        self.assertIn('全部 1 个文件均已发送成功', logs)

    def test_summary_warns_on_incomplete(self):
        """有文件没发成功时要明确警告"""
        send_list = [{'name': '张三', 'files': ['f1.txt', 'f2.txt'], 'keywords': []}]
        self._run(send_list, {
            '张三': [SendResult.permanent_fail(Stage.FILES, '失败')]})
        logs = self._logs()
        self.assertIn('有文件未成功发送', logs)

    def test_retry_success_counts_once(self):
        """重试后成功，文件只算一次"""
        send_list = [{'name': '张三', 'files': ['f1.txt'], 'keywords': []}]
        self._run(send_list, {
            '张三': [SendResult.safe_retry(Stage.SEARCH, '失败'),
                     SendResult.success(Stage.FILES)]}, max_retries=2)
        s = state['stats']
        self.assertEqual(s['files_done'], 1, '重试不应重复计数')
        self.assertEqual(s['total_files'], 1)

    def test_missing_file_counts_as_failed(self):
        """文件在磁盘上不存在 → 记为失败"""
        send_list = [{'name': '张三', 'files': ['不存在.txt'], 'keywords': []}]
        self._run(send_list, {})
        s = state['stats']
        self.assertEqual(s['files_failed'], 1)

    def test_stats_backward_compatible(self):
        """旧的联系人维度字段不能丢（前端/测试都依赖）"""
        send_list = [{'name': '张三', 'files': ['f1.txt'], 'keywords': []}]
        self._run(send_list, {})
        s = state['stats']
        for f in ('total', 'succeeded', 'failed', 'needs_review',
                  'retried_then_ok'):
            self.assertIn(f, s, f'缺少旧字段 {f}')

    def test_file_ranges_dont_repeat_after_failure(self):
        """
        回归测试：文件区间不能重复。
        失败的文件不推进 files_done，若用 files_done 推算区间，
        下一个条目会显示与失败条目相同的区间（曾出现 3/4 → 3/4）。
        """
        send_list = [
            {'name': 'A', 'files': ['f1.txt', 'f2.txt'], 'keywords': []},
            {'name': 'B', 'files': ['f3.txt'], 'keywords': []},   # 失败
            {'name': 'C', 'files': ['f4.txt'], 'keywords': []},
        ]
        self._run(send_list, {
            'A': [SendResult.success(Stage.FILES)],
            'B': [SendResult.permanent_fail(Stage.FILES, '失败')],
            'C': [SendResult.success(Stage.FILES)],
        })
        logs = self._logs()
        self.assertIn('文件 1-2/4', logs)
        self.assertIn('文件 3/4', logs)
        self.assertIn('文件 4/4', logs, '失败后下一个条目应接着 4/4，不能重复 3/4')
        # 确认没有重复区间
        import re
        ranges = re.findall(r'文件 (\d+(?:-\d+)?)/4', logs)
        starts = [int(r.split('-')[0]) for r in ranges]
        self.assertEqual(sorted(set(starts)), [1, 3, 4],
                         f'文件区间起点不应重复: {ranges}')


class TestFrontendAllocationUi(unittest.TestCase):
    """前端新增元素的静态检查"""

    @classmethod
    def setUpClass(cls):
        with web_app.app.test_client() as c:
            cls.html = c.get('/').get_data(as_text=True)

    def test_alloc_bar_exists(self):
        for eid in ('allocBar', 'allocTotal', 'allocAssigned',
                    'allocUnmatched', 'allocContacts', 'allocCoverage',
                    'allocCovFill'):
            self.assertIn(f'id="{eid}"', self.html, f'缺少 {eid}')

    def test_unmatched_box_exists(self):
        for eid in ('unmatchedBox', 'unmatchedList', 'umToggle'):
            self.assertIn(f'id="{eid}"', self.html, f'缺少 {eid}')

    def test_stats_conclusion_exists(self):
        self.assertIn('id="statsConclusion"', self.html)

    def test_new_functions_defined(self):
        for fn in ('renderAllocStats', 'toggleUnmatched'):
            self.assertIn(f'function {fn}(', self.html, f'缺少函数 {fn}')

    def test_render_preview_calls_alloc_stats(self):
        """previewMatches 必须调用 renderAllocStats"""
        js = self.html.split('<script>')[1].split('</script>')[0]
        fn = js.split('async function previewMatches()')[1].split('\nasync function')[0]
        self.assertIn('renderAllocStats', fn)

    def test_render_stats_uses_file_counts(self):
        js = self.html.split('<script>')[1].split('</script>')[0]
        fn = js.split('function renderStats(')[1].split('\n}')[0]
        self.assertIn('files_done', fn)
        self.assertIn('total_files', fn)
        self.assertIn('statsConclusion', fn)

    def test_css_classes_present(self):
        for cls in ('.alloc-bar', '.alloc-value', '.cov-fill',
                    '.unmatched-box', '.unmatched-item', '.stats-conclusion'):
            self.assertIn(cls, self.html, f'缺少样式 {cls}')

    def test_unmatched_list_defaults_to_expanded(self):
        """未匹配清单默认展开——漏掉的文件最需要被看见，不该再藏一次"""
        self.assertIn('id="unmatchedList" style="display:flex;"', self.html,
                      'unmatchedList 初始应为展开状态')
        self.assertIn('id="umToggle">收起 ▴', self.html,
                      'toggle 初始文字应为「收起」')

    def test_render_alloc_stats_expands_list(self):
        """renderAllocStats 在有未匹配时应把列表设为展开"""
        js = self.html.split('<script>')[1].split('</script>')[0]
        fn = js.split('function renderAllocStats(')[1].split('\n}')[0]
        self.assertIn("list.style.display = 'flex'", fn,
                      '渲染时应展开未匹配列表')
        self.assertIn("'收起 ▴'", fn)

    def test_summary_log_style(self):
        self.assertIn(".log-line.summary", self.html)


if __name__ == '__main__':
    unittest.main(verbosity=2)
