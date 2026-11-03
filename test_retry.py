#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重试策略单元测试
================
纯逻辑测试，**不需要微信、不需要 pyautogui / pywin32**。
直接运行： python test_retry.py
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from send_outcome import SendOutcome, SendResult, Stage
from retry import (
    RetryPolicy, RetryQueue, QueueEntry, EntryStatus,
    backoff_delays,
)


def no_sleep(_):
    """替代 time.sleep，让测试瞬间跑完并记录等待时长"""
    return None


class TestSendResult(unittest.TestCase):
    """四态定义本身的语义约束"""

    def test_success_is_ok(self):
        self.assertTrue(SendResult.success().ok)

    def test_only_safe_to_retry_allows_auto_retry(self):
        """核心安全约束：只有「确定没发出去」才允许自动重试"""
        self.assertTrue(SendResult.safe_retry(Stage.SEARCH).can_retry_automatically)
        self.assertFalse(SendResult.uncertain(Stage.FILES).can_retry_automatically)
        self.assertFalse(SendResult.permanent_fail(Stage.FILES).can_retry_automatically)
        self.assertFalse(SendResult.success().can_retry_automatically)

    def test_enter_pressed_flag(self):
        """回车是否已按下 —— 整个重试策略的安全分界线"""
        self.assertFalse(SendResult.safe_retry(Stage.FILES).enter_pressed)
        self.assertTrue(SendResult.uncertain(Stage.FILES).enter_pressed)
        self.assertTrue(SendResult.success(Stage.FILES).enter_pressed)

    def test_to_dict_serializable(self):
        d = SendResult.safe_retry(Stage.SEARCH, '超时').to_dict()
        self.assertEqual(d['outcome'], 'safe_to_retry')
        self.assertEqual(d['stage'], '搜索联系人')
        self.assertEqual(d['detail'], '超时')


class TestBackoff(unittest.TestCase):

    def test_exponential_growth(self):
        self.assertEqual(backoff_delays(3), [1.0, 3.0])
        self.assertEqual(backoff_delays(4), [1.0, 3.0, 9.0])

    def test_no_delay_when_single_attempt(self):
        self.assertEqual(backoff_delays(1), [])

    def test_zero_attempts(self):
        self.assertEqual(backoff_delays(0), [])


class TestRetryPolicy(unittest.TestCase):

    def setUp(self):
        self.policy = RetryPolicy(max_attempts=2)

    def test_rejects_invalid_max_attempts(self):
        with self.assertRaises(ValueError):
            RetryPolicy(max_attempts=0)

    def test_safe_failure_triggers_retry(self):
        e = QueueEntry(name='张三', max_attempts=2)
        self.assertEqual(
            self.policy.apply(e, SendResult.safe_retry(Stage.SEARCH)),
            'retry')
        self.assertIs(e.status, EntryStatus.RETRYING)

    def test_uncertain_never_auto_retries(self):
        """
        本测试是整个模块最重要的一条：
        回车已按下、状态不明时，绝不能自动重试，否则会给对方发两份文件。
        """
        e = QueueEntry(name='张三', max_attempts=5)
        action = self.policy.apply(e, SendResult.uncertain(Stage.FILES, '回车后超时'))
        self.assertEqual(action, 'review')
        self.assertIs(e.status, EntryStatus.NEEDS_REVIEW)
        self.assertEqual(e.attempt, 1, '不确定的发送不应消耗多次尝试')

    def test_uncertain_respects_attempt_budget(self):
        """即便开了 auto_retry_uncertain，也受尝试次数上限约束"""
        p = RetryPolicy(max_attempts=1, auto_retry_uncertain=True)
        e = QueueEntry(name='张三', max_attempts=1)
        self.assertEqual(p.apply(e, SendResult.uncertain(Stage.FILES)), 'review')

    def test_permanent_fail_skips_retry(self):
        e = QueueEntry(name='张三', max_attempts=3)
        action = self.policy.apply(e, SendResult.permanent_fail(Stage.FILES, '文件不存在'))
        self.assertEqual(action, 'failed')
        self.assertIs(e.status, EntryStatus.FAILED)

    def test_retry_stops_at_max_attempts(self):
        """尝试次数耗尽后不再重试，转failed"""
        e = QueueEntry(name='张三', max_attempts=2)
        self.policy.apply(e, SendResult.safe_retry(Stage.SEARCH))
        self.assertIs(e.status, EntryStatus.RETRYING)
        action = self.policy.apply(e, SendResult.safe_retry(Stage.SEARCH))
        self.assertEqual(action, 'failed')
        self.assertIs(e.status, EntryStatus.FAILED)
        self.assertEqual(e.attempt, 2)

    def test_success_marks_succeeded(self):
        e = QueueEntry(name='张三', max_attempts=2)
        self.assertEqual(self.policy.apply(e, SendResult.success()), 'success')
        self.assertIs(e.status, EntryStatus.SUCCEEDED)


class TestRetryQueue(unittest.TestCase):

    def test_retry_then_success(self):
        """最常见的路径：第一次失败，重试后成功"""
        q = RetryQueue(policy=RetryPolicy(max_attempts=2))
        e = q.add('张三', ['工资单.pdf'])

        q.handle(e, SendResult.safe_retry(Stage.SEARCH, '窗口未响应'), no_sleep)
        self.assertIs(e.status, EntryStatus.RETRYING)
        self.assertIn(e, q.pending())

        q.handle(e, SendResult.success(), no_sleep)
        self.assertIs(e.status, EntryStatus.SUCCEEDED)
        self.assertNotIn(e, q.pending())

        s = q.stats()
        self.assertEqual(s['succeeded'], 1)
        self.assertEqual(s['retried_then_ok'], 1, '应识别出「重试后成功」')

    def test_uncertain_goes_to_review_not_pending(self):
        q = RetryQueue(policy=RetryPolicy(max_attempts=3))
        e = q.add('李四', ['合同.pdf'])
        q.handle(e, SendResult.uncertain(Stage.FILES, '回车后无响应'), no_sleep)

        self.assertNotIn(e, q.pending(), '不确定的条目不能自动回到待发队列')
        self.assertIn(e, q.needs_review())
        self.assertEqual(q.stats()['needs_review'], 1)

    def test_confirm_sent_clears_review(self):
        """人工确认已发送 → 标记成功，移出复核队列"""
        q = RetryQueue()
        e = q.add('李四', ['合同.pdf'])
        q.handle(e, SendResult.uncertain(Stage.FILES), no_sleep)

        self.assertTrue(q.confirm_sent('李四'))
        self.assertIs(e.status, EntryStatus.SUCCEEDED)
        self.assertEqual(q.needs_review(), [])

    def test_resend_resets_attempts(self):
        """人工判定没发出去 → 重置计数重新入队，且给满重试次数"""
        q = RetryQueue(policy=RetryPolicy(max_attempts=2))
        e = q.add('李四', ['合同.pdf'])
        q.handle(e, SendResult.safe_retry(Stage.SEARCH), no_sleep)
        q.handle(e, SendResult.safe_retry(Stage.SEARCH), no_sleep)
        self.assertIs(e.status, EntryStatus.FAILED)

        e.status = EntryStatus.NEEDS_REVIEW
        self.assertTrue(q.resend('李四'))
        self.assertIs(e.status, EntryStatus.PENDING)
        self.assertEqual(e.attempt, 0, '重发应重置尝试计数')
        self.assertEqual(e.attempts_left, 2)
        self.assertIn(e, q.pending())

    def test_review_ops_on_unknown_name(self):
        q = RetryQueue()
        self.assertFalse(q.confirm_sent('查无此人'))
        self.assertFalse(q.resend('查无此人'))

    def test_backoff_actually_sleeps(self):
        """验证退避时长确实被使用"""
        slept = []
        q = RetryQueue(policy=RetryPolicy(max_attempts=3, base_delay=1.0, factor=3.0))
        e = q.add('张三', ['a.pdf'])
        q.handle(e, SendResult.safe_retry(Stage.SEARCH), slept.append)
        self.assertEqual(slept, [1.0])
        q.handle(e, SendResult.safe_retry(Stage.SEARCH), slept.append)
        self.assertEqual(slept, [1.0, 3.0])

    def test_no_sleep_on_success_or_review(self):
        slept = []
        q = RetryQueue(policy=RetryPolicy(max_attempts=3))
        e = q.add('张三', ['a.pdf'])
        q.handle(e, SendResult.success(), slept.append)
        q.handle(e, SendResult.uncertain(Stage.FILES), slept.append)
        self.assertEqual(slept, [], '成功/待复核都不该等待')

    def test_mixed_batch_stats(self):
        """混合场景：成功 / 重试后成功 / 待复核 / 彻底失败"""
        q = RetryQueue(policy=RetryPolicy(max_attempts=2))

        a = q.add('A', ['a.pdf']);q.handle(a, SendResult.success(), no_sleep)
        b = q.add('B', ['b.pdf'])
        q.handle(b, SendResult.safe_retry(Stage.SEARCH), no_sleep)
        q.handle(b, SendResult.success(), no_sleep)
        c = q.add('C', ['c.pdf']); q.handle(c, SendResult.uncertain(Stage.FILES), no_sleep)
        d = q.add('D', ['d.pdf'])
        q.handle(d, SendResult.permanent_fail(Stage.FILES, '文件不存在'), no_sleep)

        s = q.stats()
        self.assertEqual(s['total'], 4)
        self.assertEqual(s['succeeded'], 2)
        self.assertEqual(s['retried_then_ok'], 1)
        self.assertEqual(s['needs_review'], 1)
        self.assertEqual(s['failed'], 1)

    def test_history_records_every_attempt(self):
        q = RetryQueue(policy=RetryPolicy(max_attempts=2))
        e = q.add('张三', ['a.pdf'])
        q.handle(e, SendResult.safe_retry(Stage.SEARCH, '第一次'), no_sleep)
        q.handle(e, SendResult.success(), no_sleep)

        self.assertEqual(len(e.history), 2)
        self.assertEqual(e.history[0]['detail'], '第一次')
        self.assertEqual(e.history[0]['outcome'], 'safe_to_retry')
        self.assertEqual(e.history[1]['outcome'], 'success')

    def test_extend_from_send_list(self):
        """从 web_app 的 send_list 结构批量导入"""
        q = RetryQueue()
        send_list = [
            {'name': '张三', 'files': ['a.pdf', 'b.pdf'], 'keywords': ['月报', '工资']},
            {'name': '李四', 'files': ['c.pdf'], 'keywords': ['合同']},
        ]
        q.extend(send_list)
        self.assertEqual(len(q.entries), 2)
        self.assertEqual(q.entries[0].files, ['a.pdf', 'b.pdf'])
        self.assertEqual(q.entries[0].keywords, ['月报', '工资'])

    def test_extend_with_message_builder(self):
        q = RetryQueue()
        q.extend(
            [{'name': '张三', 'files': ['a.pdf']}],
            message_for=lambda n: f'你好 {n}',
        )
        self.assertEqual(q.entries[0].message, '你好 张三')

    def test_attempts_left(self):
        e = QueueEntry(name='张三', max_attempts=3)
        self.assertEqual(e.attempts_left, 3)
        e.attempt = 2
        self.assertEqual(e.attempts_left, 1)
        e.attempt = 99
        self.assertEqual(e.attempts_left, 0, '不能为负')

    def test_to_dict_serializable(self):
        import json
        q = RetryQueue()
        e = q.add('张三', ['a.pdf'], ['月报'])
        q.handle(e, SendResult.uncertain(Stage.FILES, '未知'), no_sleep)
        json.dumps(q.to_list())  # 不应抛异常


if __name__ == '__main__':
    unittest.main(verbosity=2)
