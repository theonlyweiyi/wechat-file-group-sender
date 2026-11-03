#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发送结果四态（独立模块）
========================
从 wechat_sender.py 抽出的纯定义部分，**不 import 任何微信/Windows 依赖**。

这样retry.py 与测试都能在没有 pyautogui / pywin32 的环境里运行。
wechat_sender.py 从这里 import 保持唯一真相来源。
"""

from dataclasses import dataclass
from enum import Enum


class SendOutcome(str, Enum):
    """
    一次发送尝试的结果分类。

    为什么需要四态而不是 True/False：
    键鼠模拟无法可靠确认「回车」是否真的把消息发出去了。若把所有失败
    都当作「可重试」，一旦回车其实生效了、只是后续步骤超时，重试就会造成
    **重复发送**——比失败更糟。因此以「回车是否已按下」为安全分界线：

        SAFE_TO_RETRY  失败发生在回车之前 → 确定没发出去 → 可放心重试
        UNCERTAIN      回车已按下但无法确认 → 可能已发出 → 绝不自动重试
        PERMANENT_FAIL 环境性问题（文件不存在/剪贴板被占用）→ 重试无意义
        SUCCESS        回车已按下且后续无异常
    """
    SUCCESS = 'success'
    SAFE_TO_RETRY = 'safe_to_retry'
    UNCERTAIN = 'uncertain'
    PERMANENT_FAIL = 'permanent_fail'


class Stage(str, Enum):
    """发送阶段 —— 用于日志定位失败点。"""
    CONNECT = '连接微信'
    SEARCH = '搜索联系人'
    VERIFY = '校验身份'
    MESSAGE = '发送附言'
    FILES = '发送文件'
    DONE = '完成'


@dataclass
class SendResult:
    """一次发送尝试的结构化结果。"""
    outcome: SendOutcome
    stage: Stage = Stage.DONE
    detail: str = ''
    # 关键标志：回车键是否已按下。True 表示消息可能已离开输入框。
    enter_pressed: bool = False

    @property
    def ok(self) -> bool:
        return self.outcome is SendOutcome.SUCCESS

    @property
    def can_retry_automatically(self) -> bool:
        """只有「确定没发出去」才允许自动重试。"""
        return self.outcome is SendOutcome.SAFE_TO_RETRY

    @classmethod
    def success(cls, stage=Stage.DONE, detail=''):
        return cls(SendOutcome.SUCCESS, stage, detail, enter_pressed=True)

    @classmethod
    def safe_retry(cls, stage, detail=''):
        return cls(SendOutcome.SAFE_TO_RETRY, stage, detail, enter_pressed=False)

    @classmethod
    def uncertain(cls, stage, detail=''):
        return cls(SendOutcome.UNCERTAIN, stage, detail, enter_pressed=True)

    @classmethod
    def permanent_fail(cls, stage, detail=''):
        return cls(SendOutcome.PERMANENT_FAIL, stage, detail, enter_pressed=False)

    def to_dict(self):
        return {
            'outcome': self.outcome.value,
            'stage': self.stage.value,
            'detail': self.detail,
            'enter_pressed': self.enter_pressed,
        }
