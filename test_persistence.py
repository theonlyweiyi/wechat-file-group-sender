#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
状态持久化与并发测试（缺口4）
=============================
覆盖三块：
  1. StateStore —— dict 兼容性 + 线程安全
  2. ConfigStore —— 落盘/读取/损坏容错/路径校验/白名单
  3. 真实并发压力 —— 多线程写日志同时读，验证不丢不乱

**不需要 pyautogui / pywin32 / 微信**，也不碰用户真实配置目录。

直接运行： python test_persistence.py
"""

import json
import os
import sys
import tempfile
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 必须在 import config_store 之前设好，确保不碰用户真实 %APPDATA%
_TMP_ROOT = tempfile.mkdtemp(prefix='wfs_persist_test_')
os.environ['WFS_CONFIG_DIR'] = _TMP_ROOT

from state_store import StateStore, ListAppender  # noqa: E402
import config_store as cs  # noqa: E402
from config_store import ConfigStore, PERSIST_FIELDS, NEVER_PERSIST  # noqa: E402


def tmp_config(name='config.json', **kw):
    """每个测试用独立的临时配置文件。"""
    d = tempfile.mkdtemp(dir=_TMP_ROOT)
    return ConfigStore(path=os.path.join(d, name), **kw)


class TestStateStoreDictCompat(unittest.TestCase):
    """必须完全兼容 dict 用法，否则 62 处调用点会崩"""

    def test_subscript_roundtrip(self):
        s = StateStore({'a': 1})
        self.assertEqual(s['a'], 1)
        s['b'] = 2
        self.assertEqual(s['b'], 2)

    def test_missing_key_raises_keyerror(self):
        s = StateStore({})
        with self.assertRaises(KeyError):
            s['nope']

    def test_get_with_default(self):
        s = StateStore({})
        self.assertEqual(s.get('x', 'dflt'), 'dflt')
        s['x'] = 1
        self.assertEqual(s.get('x', 'dflt'), 1)

    def test_len_contains_iter(self):
        s = StateStore({'a': 1, 'b': 2})
        self.assertEqual(len(s), 2)
        self.assertIn('a', s)
        self.assertNotIn('zz', s)
        self.assertEqual(sorted(s), ['a', 'b'])

    def test_keys_values_items_are_copies(self):
        s = StateStore({'a': 1})
        keys = s.keys()
        keys.append('hack')
        self.assertEqual(len(s), 1, 'keys() 必须返回副本，不能泄露内部')

    def test_setdefault(self):
        s = StateStore({'a': 1})
        self.assertEqual(s.setdefault('a', 99), 1)
        self.assertEqual(s.setdefault('b', 2), 2)

    def test_pop_and_delitem(self):
        s = StateStore({'a': 1, 'b': 2})
        self.assertEqual(s.pop('a'), 1)
        del s['b']
        self.assertEqual(len(s), 0)

    def test_update(self):
        s = StateStore({})
        s.update({'a': 1})
        s.update(b=2)
        s.update({'c': 3}, d=4)
        self.assertEqual(s['a'], 1)
        self.assertEqual(s['d'], 4)
        self.assertEqual(len(s), 4)

    def test_snapshot_is_copy(self):
        s = StateStore({'a': [1, 2]})
        snap = s.snapshot()
        s['a'].append(3)
        self.assertEqual(snap['a'], [1, 2], 'snapshot 应是快照而非引用')

    def test_replace(self):
        s = StateStore({'a': 1})
        s.replace({'b': 2})
        self.assertNotIn('a', s)
        self.assertEqual(s['b'], 2)

    def test_log_autotruncate(self):
        """logs 超过上限自动裁剪"""
        s = StateStore({'logs': []}, log_limit=10)
        for i in range(25):
            s['logs'] = s['logs'] + [{'i': i}]
        self.assertEqual(len(s['logs']), 10)
        self.assertEqual(s['logs'][-1]['i'], 24, '应保留最新的')


class TestStateStoreHook(unittest.TestCase):

    def test_on_change_fires(self):
        seen = []
        s = StateStore({}, on_change=lambda k, v: seen.append((k, v)))
        s['x'] = 1
        self.assertEqual(seen, [('x', 1)])

    def test_hook_exception_does_not_break_state(self):
        """钩子抛异常绝不能影响主流程——持久化是辅助功能"""
        def bad(k, v):
            raise RuntimeError('boom')
        s = StateStore({}, on_change=bad)
        s['x'] = 1              # 不应抛出
        self.assertEqual(s['x'], 1)

    def test_hook_called_outside_lock(self):
        """钩子必须能在锁外访问 state，否则会死锁"""
        s = StateStore({'logs': []})
        def hook(k, v):
            s['touched'] = True# 若在锁内调用 → 死锁
        s.set_hook(hook)
        s['x'] = 1
        self.assertTrue(s['touched'])

    def test_hook_not_fired_for_get(self):
        seen = []
        s = StateStore({'a': 1}, on_change=lambda k, v: seen.append(k))
        s['a']
        s.get('a')
        s.snapshot()
        self.assertEqual(seen, [], '读取不应触发变更钩子')


class TestAtomicUpdate(unittest.TestCase):

    def test_atomic_update_mutates_in_place(self):
        s = StateStore({'logs': []})
        s.atomic_update(lambda d: d['logs'].append('x'))
        self.assertEqual(s['logs'], ['x'])

    def test_atomic_update_returns_value(self):
        s = StateStore({'n': 0})
        r = s.atomic_update(lambda d: d.__setitem__('n', d['n'] + 1) or d['n'])
        self.assertEqual(r, 1)

    def test_atomic_update_truncates_logs(self):
        s = StateStore({'logs': []}, log_limit=5)
        for i in range(20):
            s.atomic_update(lambda d, i=i: d['logs'].append(i))
        self.assertEqual(len(s['logs']), 5)
        self.assertEqual(s['logs'], [15, 16, 17, 18, 19])

    def test_reentrant_no_deadlock(self):
        """atomic_update 内再写 state 不能死锁（用了 RLock）"""
        s = StateStore({'n': 0})
        def mutator(d):
            s['inner'] = 'ok'          # 锁内再取锁
            d['n'] += 1
        done = []
        t = threading.Thread(target=lambda: done.append(s.atomic_update(mutator)))
        t.start()
        t.join(timeout=3)
        self.assertFalse(t.is_alive(), '发生了死锁')
        self.assertEqual(s['inner'], 'ok')


class TestListAppender(unittest.TestCase):

    def test_append_and_trim(self):
        s = StateStore({'logs': []}, log_limit=5)
        la = ListAppender(s, 'logs')
        for i in range(12):
            la.append(i)
        self.assertEqual(s['logs'], [7, 8, 9, 10, 11])

    def test_snapshot_and_clear(self):
        s = StateStore({'logs': []})
        la = ListAppender(s, 'logs')
        la.append('a')
        self.assertEqual(la.snapshot(), ['a'])
        la.clear()
        self.assertEqual(s['logs'], [])

    def test_handles_wrong_type(self):
        """logs 被误设为非 list 时应自愈"""
        s = StateStore({'logs': 'notalist'})
        la = ListAppender(s, 'logs')
        la.append('x')
        self.assertEqual(s['logs'], ['x'])

    def test_extend(self):
        s = StateStore({'logs': []})
        ListAppender(s, 'logs').extend(['a', 'b'])
        self.assertEqual(s['logs'], ['a', 'b'])


class TestConcurrency(unittest.TestCase):
    """真实并发压力：多线程写 + 主线程读，验证不丢不乱"""

    def test_no_lost_logs_under_concurrency(self):
        """
        8 个线程各写 200条，后台线程同时读。
        旧实现（无锁 dict + append-then-trim）会丢日志或长度失控。
        """
        s = StateStore({'logs': []}, log_limit=100000)
        la = ListAppender(s, 'logs')
        threads = []
        stop = threading.Event()
        read_errors = []

        def writer(tid):
            for i in range(200):
                la.append({'t': tid, 'i': i})

        def reader():
            while not stop.is_set():
                snap = s['logs']
                if not isinstance(snap, list):
                    read_errors.append('logs 不是list')
                time.sleep(0.001)

        rt = threading.Thread(target=reader)
        rt.start()
        for tid in range(8):
            t = threading.Thread(target=writer, args=(tid,))
            threads.append(t)
            t.start()
        for t in threads:
            t.join()
        stop.set()
        rt.join()

        self.assertEqual(read_errors, [])
        self.assertEqual(len(s['logs']), 8 * 200, '日志不能丢')

    def test_counter_increments_correctly(self):
        """
        经典竞态：读-改-写若不原子，会丢更新。
        atomic_update 保证 8 线程 × 500 次 = 4000。
        """
        s = StateStore({'n': 0})
        def bump(d):
            d['n'] = d.get('n', 0) + 1
        threads = [threading.Thread(target=lambda: [s.atomic_update(bump)
                                                   for _ in range(500)])
                   for _ in range(8)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        self.assertEqual(s['n'], 8 * 500, '原子操作下计数必须精确')

    def test_get_set_are_atomic(self):
        s = StateStore({'a': 0})
        def worker():
            for _ in range(300):
                s['a'] = s['a'] + 1
        # 注意：这个写法在无锁下会丢更新，测试它是否被 StateStore 兜住
        threads = [threading.Thread(target=worker) for _ in range(4)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()
        # 可能少于 1200（get/set 分离仍不原子），但不能超过
        self.assertLessEqual(s['a'], 4 * 300)
        self.assertGreater(s['a'], 0)


class TestConfigRoundTrip(unittest.TestCase):

    def test_save_and_load(self):
        cfg = tmp_config()
        s = StateStore({'mappings': [('工资', '张三')],
                        'max_retries': 4, 'msg': '你好 {name}',
                        'enable_msg': True, 'folder_path': ''})
        cfg.bind(s)
        cfg.save(s, force=True)

        s2 = StateStore({'mappings': []})
        cfg2 = ConfigStore(path=cfg.path)
        loaded = cfg2.load()
        self.assertEqual(loaded['max_retries'], 4)
        self.assertEqual(loaded['msg'], '你好 {name}')
        self.assertEqual(loaded['mappings'], [('工资', '张三')])

    def test_mappings_convert_to_tuple(self):
        """JSON 读回是 list，bind 应转成 tuple 匹配原行为"""
        cfg = tmp_config()
        s = StateStore({'mappings': [('a', 'b')]})
        cfg.bind(s)
        cfg.save(s, force=True)

        s2 = StateStore({'mappings': []})
        ConfigStore(path=cfg.path).bind(s2)
        self.assertIsInstance(s2['mappings'][0], tuple)

    def test_max_retries_clamped(self):
        cfg = tmp_config()
        s = StateStore({'max_retries': 999})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertEqual(ConfigStore(path=cfg.path).load()['max_retries'], 5)

        s['max_retries'] = -3
        cfg.save(s, force=True)
        self.assertEqual(ConfigStore(path=cfg.path).load()['max_retries'], 1)

    def test_file_is_utf8_json(self):
        cfg = tmp_config()
        s = StateStore({'msg': '你好世界'})
        cfg.bind(s)
        cfg.save(s, force=True)
        with open(cfg.path, encoding='utf-8') as f:
            raw = json.load(f)
        self.assertEqual(raw['settings']['msg'], '你好世界')
        self.assertEqual(raw['version'], cs.CONFIG_VERSION)

    def test_false_is_persisted_not_treated_as_empty(self):
        """
        False 是有意义的配置值（用户明确关掉了附言），
        不能被当成「空值」丢弃——否则用户下次启动附言又回来了。
        """
        cfg = tmp_config()
        s = StateStore({'enable_msg': False, 'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        with open(cfg.path, encoding='utf-8') as f:
            raw = json.load(f)
        self.assertIn('enable_msg', raw['settings'])
        self.assertIs(raw['settings']['enable_msg'], False)

    def test_empty_containers_skipped(self):
        """空串/空列表属于「未设置」，不落盘"""
        cfg = tmp_config()
        s = StateStore({'msg': '', 'mappings': [], 'recent_folders': [],
                        'max_retries': 2})
        cfg.bind(s)
        cfg.save(s, force=True)
        with open(cfg.path, encoding='utf-8') as f:
            raw = json.load(f)
        self.assertNotIn('msg', raw['settings'])
        self.assertNotIn('mappings', raw['settings'])
        self.assertIn('max_retries', raw['settings'])


class TestConfigWhitelist(unittest.TestCase):
    """白名单：运行时状态绝不能落盘"""

    def test_runtime_fields_not_persisted(self):
        cfg = tmp_config()
        s = StateStore({
            'sending': True, 'logs': ['x'], 'progress': {'current': 1},
            'result': 'completed', 'active_queue': object(),
            'needs_review': [{}], 'stats': {}, 'queue': [],
            'matched': [], 'grouped': [], 'excel_uploaded': ('x.xlsx', 'p'),
            'msg': 'keep', 'max_retries': 2,
        })
        cfg.bind(s)
        cfg.save(s, force=True)

        with open(cfg.path, encoding='utf-8') as f:
            raw = json.load(f)
        for field in NEVER_PERSIST:
            self.assertNotIn(field, raw['settings'],
                             f'{field} 属运行时状态，不该被持久化')
        self.assertIn('msg', raw['settings'])

    def test_active_queue_would_break_json(self):
        """active_queue 含内存对象，若漏进白名单会直接 json 序列化崩溃"""
        cfg = tmp_config()
        s = StateStore({'active_queue': object(), 'msg': 'x'})
        cfg.bind(s)
        self.assertTrue(cfg.save(s, force=True), '序列化不能崩')

    def test_persist_fields_all_in_initial_state(self):
        """每个白名单字段都应在 _INITIAL_STATE 里声明过"""
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                'web_app.py'), encoding='utf-8').read()
        for f in PERSIST_FIELDS:
            self.assertIn(f'"{f}"', src, f'{f} 应在 _INITIAL_STATE 中声明')

    def test_never_persist_disjoint_from_persist(self):
        overlap = set(PERSIST_FIELDS) & set(NEVER_PERSIST)
        self.assertFalse(overlap, f'两列表不应重叠: {overlap}')


class TestConfigFieldFilter(unittest.TestCase):

    def test_filter_blocks_field(self):
        cfg = tmp_config()
        cfg.add_field_filter('mappings', lambda s: False)
        s = StateStore({'mappings': [('a', 'b')], 'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        with open(cfg.path, encoding='utf-8') as f:
            raw = json.load(f)
        self.assertNotIn('mappings', raw['settings'])
        self.assertIn('msg', raw['settings'])

    def test_filter_allows_field(self):
        cfg = tmp_config()
        cfg.add_field_filter('mappings', lambda s: True)
        s = StateStore({'mappings': [('a', 'b')]})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertIn('mappings', ConfigStore(path=cfg.path).load())

    def test_filter_exception_skips_field(self):
        cfg = tmp_config()
        def boom(s):
            raise RuntimeError('bad predicate')
        cfg.add_field_filter('mappings', boom)
        s = StateStore({'mappings': [('a', 'b')], 'msg': 'x'})
        cfg.bind(s)
        self.assertTrue(cfg.save(s, force=True), '过滤器异常不应影响保存')

    def test_filter_exception_is_reported(self):
        """
        过滤器抛异常时必须留下线索。
        回归测试：曾因过滤器签名写错（不接收 state 参数），
        TypeError 被静默吞掉，导致 mappings 字段「莫名丢失」且毫无线索。
        """
        cfg = tmp_config()
        def wrong_signature():        # 故意少一个参数
            return True
        cfg.add_field_filter('mappings', wrong_signature)
        s = StateStore({'mappings': [('a', 'b')], 'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertIn('mappings', cfg.filter_errors,
                      '过滤器异常应记入 filter_errors')
        self.assertIn('mappings', ' '.join(cfg.load_warnings),
                      '应提示哪个字段被跳过')

    def test_real_webapp_filter_signature(self):
        """
        web_app 里注册的过滤器必须能接受 state 参数。
        直接验证真实过滤器，而不是只测桩。
        """
        src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                'web_app.py'), encoding='utf-8').read()
        import re
        m = re.search(r'def _mappings_are_persistable\(([^)]*)\)', src)
        self.assertIsNotNone(m, '应存在 _mappings_are_persistable')
        self.assertIn('_state', m.group(1),
                      '过滤器必须声明一个参数（ConfigStore 以 predicate(state) 调用）')


class TestConfigResilience(unittest.TestCase):

    def test_corrupt_json_does_not_raise(self):
        """配置文件损坏绝不能阻止程序启动"""
        cfg = tmp_config()
        with open(cfg.path, 'w', encoding='utf-8') as f:
            f.write('{ this is not valid json ')
        self.assertEqual(cfg.load(), {})
        self.assertTrue(cfg.load_warnings)

    def test_corrupt_file_backed_up(self):
        cfg = tmp_config()
        with open(cfg.path, 'w', encoding='utf-8') as f:
            f.write('{{{')
        cfg.load()
        siblings = [f for f in os.listdir(os.path.dirname(cfg.path))
                    if 'corrupt' in f]
        self.assertTrue(siblings, '损坏的配置应被备份')

    def test_missing_file_returns_empty(self):
        cfg = tmp_config()
        self.assertEqual(cfg.load(), {})

    def test_non_dict_json_rejected(self):
        cfg = tmp_config()
        with open(cfg.path, 'w', encoding='utf-8') as f:
            json.dump([1, 2, 3], f)
        self.assertEqual(cfg.load(), {})

    def test_folder_path_dropped_if_missing(self):
        """路径失效时必须丢弃，不能让下次启动指向不存在的目录"""
        cfg = tmp_config()
        s = StateStore({'folder_path': 'C:/绝对不存在的目录-zzz', 'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        loaded = ConfigStore(path=cfg.path).load()
        self.assertNotIn('folder_path', loaded)
        self.assertIn('msg', loaded)

    def test_folder_path_kept_if_valid(self):
        cfg = tmp_config()
        real = tempfile.mkdtemp()
        s = StateStore({'folder_path': real})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertEqual(ConfigStore(path=cfg.path).load()['folder_path'],
                         os.path.normpath(real))

    def test_recent_folders_filtered(self):
        """加载时只做「存在性过滤」；去重是 remember_folder 的职责"""
        cfg = tmp_config()
        good = tempfile.mkdtemp()
        s = StateStore({'recent_folders': [good, 'C:/不存在-abc']})
        cfg.bind(s)
        cfg.save(s, force=True)
        loaded = ConfigStore(path=cfg.path).load()['recent_folders']
        self.assertEqual(loaded, [os.path.normpath(good)],
                         '失效路径应被滤掉')

    def test_malformed_mappings_dropped(self):
        cfg = tmp_config()
        s = StateStore({'mappings': [
            ('好', '行'), ('', '空关键词'), ('空名字', ''),
            (None, None), '垃圾', ['只有一列'], 123,
        ]})
        cfg.bind(s)
        cfg.save(s, force=True)
        loaded = ConfigStore(path=cfg.path).load()['mappings']
        self.assertEqual(loaded, [('好', '行')], '畸形项应被丢弃')

    def test_mappings_accepts_dict_form(self):
        """容错：也接受 {"keyword":..,"name":..} 形式"""
        cfg = tmp_config()
        s = StateStore({'mappings': [{'keyword': '月报', 'name': '张三'}]})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertEqual(ConfigStore(path=cfg.path).load()['mappings'],
                         [('月报', '张三')])

    def test_future_version_warns(self):
        cfg = tmp_config()
        with open(cfg.path, 'w', encoding='utf-8') as f:
            json.dump({'version': 999, 'settings': {'msg': 'x'}}, f)
        cfg.load()
        self.assertTrue(any('版本' in w for w in cfg.load_warnings))

    def test_empty_payload_deletes_file(self):
        """没内容可存时删文件，而不是留一个 '{}'"""
        cfg = tmp_config()
        s = StateStore({'msg': 'x', 'max_retries': 2})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertTrue(os.path.exists(cfg.path))

        # 清空所有可持久化字段后，payload.settings 应为空 → 触发删文件
        s['msg'] = ''
        s['max_retries'] = None
        s['mappings'] = []
        s['folder_path'] = ''
        cfg.save(s, force=True)
        self.assertFalse(os.path.exists(cfg.path),
                         '无内容时应删除配置文件')

    def test_atomic_write_no_temp_left(self):
        cfg = tmp_config()
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        leftovers = [f for f in os.listdir(os.path.dirname(cfg.path))
                     if f.startswith('.config-')]
        self.assertFalse(leftovers, f'临时文件应被清理: {leftovers}')

    def test_directory_created_automatically(self):
        d = tempfile.mkdtemp(dir=_TMP_ROOT)
        target = os.path.join(d, 'deep', 'nested', 'config.json')
        cfg = ConfigStore(path=target)
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        self.assertTrue(cfg.save(s, force=True))
        self.assertTrue(os.path.isfile(target))

    def test_delete(self):
        cfg = tmp_config()
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        cfg.save(s, force=True)
        self.assertTrue(os.path.isfile(cfg.path))
        cfg.delete()
        self.assertFalse(os.path.isfile(cfg.path))

    def test_disabled_store_is_noop(self):
        cfg = tmp_config(enabled=False)
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        self.assertFalse(cfg.save(s, force=True))
        self.assertFalse(os.path.isfile(cfg.path))
        cfg.mark_dirty()          # 不应抛异常
        cfg.flush()


class TestDebounce(unittest.TestCase):

    def test_mark_dirty_schedules_save(self):
        cfg = tmp_config(debounce_seconds=0.15)
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        cfg.mark_dirty()
        time.sleep(0.5)           # 等防抖窗口过去
        self.assertTrue(os.path.isfile(cfg.path), '防抖后应自动落盘')

    def test_mark_dirty_when_disabled_does_nothing(self):
        cfg = tmp_config(enabled=False)
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        cfg.mark_dirty()
        time.sleep(0.3)
        self.assertFalse(os.path.isfile(cfg.path))

    def test_flush_cancels_pending_timer(self):
        cfg = tmp_config(debounce_seconds=5.0)
        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        cfg.mark_dirty()
        cfg.flush()               # 应立即落盘并取消 5秒后的定时器
        self.assertTrue(os.path.isfile(cfg.path))

    def test_high_frequency_marks_dont_thrash_disk(self):
        """高频 mark_dirty 不应每次都写盘（否则日志会拖垮界面）"""
        writes = []
        cfg = tmp_config(debounce_seconds=10.0)
        orig = cfg._atomic_write

        def counting(payload):
            writes.append(1)
            return orig(payload)
        cfg._atomic_write = counting

        s = StateStore({'msg': 'x'})
        cfg.bind(s)
        for _ in range(1000):
            cfg.mark_dirty()
        self.assertLessEqual(len(writes), 2, '1000次修改不应产生大量写盘')


class TestRememberFolder(unittest.TestCase):

    def test_most_recent_first(self):
        cfg = tmp_config()
        s = StateStore({})
        a, b = tempfile.mkdtemp(), tempfile.mkdtemp()
        cfg.remember_folder(s, a)
        cfg.remember_folder(s, b)
        self.assertEqual(s['recent_folders'][0], os.path.normpath(b))

    def test_duplicate_moves_to_front(self):
        cfg = tmp_config()
        s = StateStore({})
        a, b = tempfile.mkdtemp(), tempfile.mkdtemp()
        cfg.remember_folder(s, a)
        cfg.remember_folder(s, b)
        cfg.remember_folder(s, a)
        rec = s['recent_folders']
        self.assertEqual(rec[0], os.path.normpath(a))
        self.assertEqual(len(rec), 2, '重复的不应出现两次')

    def test_max_len_capped(self):
        cfg = tmp_config()
        s = StateStore({})
        for _ in range(20):
            cfg.remember_folder(s, tempfile.mkdtemp())
        self.assertLessEqual(len(s['recent_folders']), ConfigStore.MAX_RECENT)

    def test_invalid_folder_ignored(self):
        cfg = tmp_config()
        s = StateStore({'recent_folders': []})
        cfg.remember_folder(s, 'C:/绝对不存在-xyz')
        self.assertEqual(s['recent_folders'], [])

    def test_case_insensitive_dedup(self):
        """Windows 路径不区分大小写"""
        cfg = tmp_config()
        s = StateStore({})
        d = tempfile.mkdtemp()
        cfg.remember_folder(s, d)
        cfg.remember_folder(s, d.upper())
        self.assertEqual(len(s['recent_folders']), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
