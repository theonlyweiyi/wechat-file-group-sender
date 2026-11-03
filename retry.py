#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重试队列
========
决定「某个联系人失败了该怎么办」。

核心规则（与 wechat_sender.SendOutcome 配套）
------------------------------------------------
    SUCCESS         成功，结束
    SAFE_TO_RETRY   确定没发出去 → 指数退避后自动重试
    UNCERTAIN       可能已发出   → **绝不自动重试**，进人工复核队列
    PERMANENT_FAIL  重试无意义（文件不存在/剪贴板被占/身份不匹配）→ 直接记账

为什么 UNCERTAIN 不能自动重试
------------------------------
键鼠模拟无法确认「回车」是否真的让消息离开了输入框。如果回车其实生效了、
只是紧随其后的步骤超时，此时重试会给对方发两份文件——比失败严重得多。
所以状态不明时宁可交给人判断。

本模块是纯逻辑，不 import 微信相关模块，可独立单元测试。
"""

import time
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional

from send_outcome import SendOutcome, SendResult, Stage


# ══════════════════════════════════════════════════════════════
#  队列条目
# ══════════════════════════════════════════════════════════════

class EntryStatus(str, Enum):
    PENDING = 'pending'          # 待发送
    RETRYING = 'retrying'        # 重试中
    SUCCEEDED = 'succeeded'      # 已成功（含重试后成功）
    NEEDS_REVIEW = 'needs_review'  # ⚠ 状态不明，等人工判断
    FAILED = 'failed'            # 彻底失败（含重试耗尽）


@dataclass
class QueueEntry:
    """发送队列中的一条记录（通常= 一个联系人 + 其文件）。"""
    name: str                              # 联系人
    files: List[str] = field(default_factory=list)
    keywords: List[str] = field(default_factory=list)
    message: Optional[str] = None
    status: EntryStatus = EntryStatus.PENDING

    attempt: int = 0                       # 已尝试次数
    max_attempts: int = 2                  # 最多尝试次数（含首次）
    last_result: Optional[dict] = None     # 最后一次失败的 SendResult.to_dict()
    history: List[dict] = field(default_factory=list)  # 每次尝试的记录

    # ── 便捷属性 ──
    @property
    def attempts_left(self) -> int:
        return max(0, self.max_attempts - self.attempt)

    @property
    def display_files(self) -> str:
        return ", ".join(self.files) if self.files else "(无文件)"

    def can_auto_retry(self, result: SendResult) -> bool:
        """
        能否自动重试。两个条件必须同时满足：
          1) 结果是 SAFE_TO_RETRY（确定没发出去）
          2) 还有剩余尝试次数
        """
        return result.can_retry_automatically and self.attempts_left > 0

    def record(self, result: SendResult, ts: float = None):
        """记录一次尝试结果。"""
        self.attempt += 1
        rec = {
            'attempt': self.attempt,
            'outcome': result.outcome.value,
            'stage': result.stage.value,
            'detail': result.detail,
            'enter_pressed': result.enter_pressed,
            'ts': ts if ts is not None else time.time(),
        }
        self.history.append(rec)
        self.last_result = result.to_dict()

    def to_dict(self):
        return {
            'name': self.name,
            'files': self.files,
            'keywords': self.keywords,
            'message': self.message,
            'status': self.status.value,
            'attempt': self.attempt,
            'max_attempts': self.max_attempts,
            'attempts_left': self.attempts_left,
            'last_result': self.last_result,
            'history': self.history,
        }


# ══════════════════════════════════════════════════════════════
#  重试策略
# ══════════════════════════════════════════════════════════════

def backoff_delays(max_attempts: int, base: float = 1.0, factor: float = 3.0) -> List[float]:
    """
    指数退避间隔。max_attempts=3, base=1, factor=3 → [1.0, 3.0]
    （首次尝试不等待，只有重试才等）
    """
    return [base * (factor ** i) for i in range(max(0, max_attempts - 1))]


class RetryPolicy:
    """
    重试策略。纯决策，不执行动作，便于测试。

    用法：
        policy = RetryPolicy(max_attempts=2)
        if policy.should_retry(entry, result):
            delay = policy.delay_for(entry)
            time.sleep(delay)
    """

    def __init__(self, max_attempts: int = 2, base_delay: float = 1.0,
                 factor: float = 3.0, auto_retry_uncertain: bool = False):
        """
        max_attempts     每个联系人最多尝试几次（含首次），1 = 不重试
        base_delay       首次重试前等待秒数
        factor           退避倍数
        auto_retry_uncertain
                        是否对 UNCERTAIN 自动重试。**默认 False，请勿轻易改 True**：
                        打开它可能导致给同一联系人重复发送文件。
        """
        if max_attempts < 1:
            raise ValueError('max_attempts 必须 >= 1')
        self.max_attempts = max_attempts
        self.base_delay = base_delay
        self.factor = factor
        self.auto_retry_uncertain = auto_retry_uncertain

    def should_retry(self, entry: QueueEntry, result: SendResult) -> bool:
        if self.auto_retry_uncertain and result.outcome is SendOutcome.UNCERTAIN:
            return entry.attempts_left > 0
        return entry.can_auto_retry(result)

    def delay_for(self, entry: QueueEntry) -> float:
        """第 attempt 次失败后、下次尝试前的等待秒数。"""
        delays = backoff_delays(self.max_attempts, self.base_delay, self.factor)
        idx = min(max(0, entry.attempt - 1), len(delays) - 1) if delays else 0
        return delays[idx] if delays else 0.0

    def apply(self, entry: QueueEntry, result: SendResult) -> str:
        """
        根据结果更新 entry.status，返回本次失败的处理动作描述。

        返回值之一：'success' / 'retry' / 'review' / 'failed'
        """
        entry.record(result)

        if result.outcome is SendOutcome.SUCCESS:
            entry.status = EntryStatus.SUCCEEDED
            return 'success'

        if self.should_retry(entry, result):
            entry.status = EntryStatus.RETRYING
            return 'retry'

        if result.outcome is SendOutcome.UNCERTAIN:
            # 回车已按下、无法确认 → 交人工，绝不自动重试
            entry.status = EntryStatus.NEEDS_REVIEW
            return 'review'

        entry.status = EntryStatus.FAILED
        return 'failed'


# ══════════════════════════════════════════════════════════════
#  队列管理
# ══════════════════════════════════════════════════════════════

class RetryQueue:
    """
    管理待发列表 + 复核队列。

    典型用法见 _send_worker：
        q = RetryQueue(policy=RetryPolicy(max_attempts=2))
        q.add(name, files, keywords, message)
        for entry in q.pending():
            result = sender.send(...)
            action = q.handle(entry, result, sleep_fn=time.sleep)
    """

    def __init__(self, policy: Optional[RetryPolicy] = None):
        self.policy = policy or RetryPolicy()
        self.entries: List[QueueEntry] = []

    # ── 构建 ──
    def add(self, name, files, keywords=None, message=None) -> QueueEntry:
        e = QueueEntry(
            name=name,
            files=list(files),
            keywords=list(keywords or []),
            message=message,
            max_attempts=self.policy.max_attempts,
        )
        self.entries.append(e)
        return e

    def extend(self, send_list, message_for=None):
        """
        从 _send_worker 的 send_list 批量导入。
        send_list: [{'name':..., 'files':[...], 'keywords':[...]}, ...]
        message_for: 可选 callable(name) -> str|None
        """
        for m in send_list:
            msg = message_for(m['name']) if message_for else None
            self.add(m['name'], m.get('files', []), m.get('keywords', []), msg)

    # ── 消费 ──
    def pending(self) -> List[QueueEntry]:
        return [e for e in self.entries
                if e.status in (EntryStatus.PENDING, EntryStatus.RETRYING)]

    def handle(self, entry: QueueEntry, result: SendResult,
               sleep_fn=time.sleep) -> str:
        """
        应用策略并按需退避等待。返回动作。

        sleep_fn 传 None 则只决策不等待——调用方若想自己控制退避节奏
        （如在重试循环里先打日志再等待），用这个形式。
        """
        action = self.policy.apply(entry, result)
        if action == 'retry' and sleep_fn is not None:
            sleep_fn(self.policy.delay_for(entry))
        return action

    # ── 人工复核队列 ──
    def needs_review(self) -> List[QueueEntry]:
        return [e for e in self.entries if e.status is EntryStatus.NEEDS_REVIEW]

    def confirm_sent(self, name: str) -> bool:
        """人工确认「实际已发送」→ 标记成功，移出复核队列。"""
        for e in self.needs_review():
            if e.name == name:
                e.status = EntryStatus.SUCCEEDED
                return True
        return False

    def resend(self, name: str) -> bool:
        """人工判定「实际没发出去」→ 重置为待发送，重新入队。"""
        for e in self.needs_review():
            if e.name == name:
                e.status = EntryStatus.PENDING
                e.attempt = 0            # 重置计数，给满重试次数
                e.max_attempts = self.policy.max_attempts
                return True
        return False

    # ── 统计 ──
    def stats(self) -> dict:
        c = {}
        for e in self.entries:
            c[e.status.value] = c.get(e.status.value, 0) + 1
        return {
            'total': len(self.entries),
            'succeeded': c.get('succeeded', 0),
            'failed': c.get('failed', 0),
            'needs_review': c.get('needs_review', 0),
            'retried_then_ok': sum(
                1 for e in self.entries
                if e.status is EntryStatus.SUCCEEDED and e.attempt > 1
            ),
            'by_status': c,
        }

    def to_list(self) -> List[dict]:
        return [e.to_dict() for e in self.entries]
