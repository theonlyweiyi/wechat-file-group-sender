#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试文件匹配逻辑"""

import os
import sys
import tempfile
import shutil

# 添加主程序路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from wechat_file_sender import WeChatFileSenderApp
import tkinter as tk


def test_match_logic():
    """测试关键词匹配逻辑"""
    print("=" * 60)
    print("测试: 文件匹配逻辑")
    print("=" * 60)

    # 创建临时测试目录
    test_dir = tempfile.mkdtemp(prefix="wxtest_")
    print(f"\n测试目录: {test_dir}")

    # 创建模拟文件
    test_files = [
        "8月月报.xlsx",
        "7月月报.pdf",
        "2024年08月工资单.xlsx",
        "劳动合同-张三.pdf",
        "增值税发票_001.pdf",
        "增值税发票_002.pdf",
        " unrelated_file.txt",  # 不匹配任何关键词
        "README.md",
    ]
    for fname in test_files:
        fpath = os.path.join(test_dir, fname)
        with open(fpath, "w") as f:
            f.write("test content")

    # 创建映射
    mappings = [
        ("月报", "张三"),
        ("工资单", "李四"),
        ("合同", "王五"),
        ("发票", "赵六"),
        ("不存在的关键词", "测试用户"),  # 不应该匹配到任何文件
    ]

    # 使用 App 的方法进行测试（不启动GUI）
    app = WeChatFileSenderApp.__new__(WeChatFileSenderApp)

    # 扫描文件夹
    files = app._scan_folder(test_dir)
    print(f"\n扫描到 {len(files)} 个文件:")
    for fname, _ in files:
        print(f"  - {fname}")

    # 匹配
    matched = app._match_files(mappings, files)

    print(f"\n匹配结果:")
    all_passed = True
    for keyword, name, file_list in matched:
        status = f"{len(file_list)} 个文件" if file_list else "无匹配"
        print(f"  关键词「{keyword}」→ {name}: {status}")
        for f in file_list:
            print(f"    └─ {os.path.basename(f)}")

    # 验证
    print("\n验证:")
    checks = [
        ("月报", 2, "应匹配到 8月月报.xlsx 和 7月月报.pdf"),
        ("工资单", 1, "应匹配到 2024年08月工资单.xlsx"),
        ("合同", 1, "应匹配到 劳动合同-张三.pdf"),
        ("发票", 2, "应匹配到两个发票文件"),
        ("不存在的关键词", 0, "不应匹配到任何文件"),
    ]

    for keyword, expected, desc in checks:
        actual = next(len(f) for k, n, f in matched if k == keyword)
        passed = actual == expected
        symbol = "✅" if passed else "❌"
        print(f"  {symbol} {keyword}: 期望 {expected}, 实际 {actual} — {desc}")
        if not passed:
            all_passed = False

    # 测试大小写不敏感
    print("\n测试大小写不敏感:")
    test_files_upper = ["REPORT_August.xlsx", "report_july.pdf"]
    for fname in test_files_upper:
        fpath = os.path.join(test_dir, fname)
        with open(fpath, "w") as f:
            f.write("test")
    files2 = app._scan_folder(test_dir)
    matched2 = app._match_files([("report", "测试")], files2)
    case_insensitive_ok = len(matched2[0][2]) == 2
    symbol = "✅" if case_insensitive_ok else "❌"
    print(f"  {symbol} 关键词「report」匹配大小写混合文件名: {len(matched2[0][2])} 个文件")
    if not case_insensitive_ok:
        all_passed = False

    # 清理
    shutil.rmtree(test_dir)

    print(f"\n{'=' * 60}")
    print(f"结果: {'全部通过 ✅' if all_passed else '有失败项 ❌'}")
    print(f"{'=' * 60}")
    return all_passed


def test_excel_reading():
    """测试 Excel 读取"""
    print("\n" + "=" * 60)
    print("测试: Excel 映射表读取")
    print("=" * 60)

    excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "发送映射模板.xlsx")
    if not os.path.exists(excel_path):
        print(f"❌ 模板文件不存在: {excel_path}")
        return False

    app = WeChatFileSenderApp.__new__(WeChatFileSenderApp)
    mappings = app._read_excel_mapping(excel_path)

    print(f"\n读取到 {len(mappings)} 条映射:")
    for keyword, name in mappings:
        print(f"  关键词: {keyword:<10} → 微信昵称: {name}")

    passed = len(mappings) > 0
    symbol = "✅" if passed else "❌"
    print(f"\n{symbol} Excel 读取: {'成功' if passed else '失败'}")
    return passed


if __name__ == "__main__":
    # 不启动GUI，只测试逻辑
    # 创建一个隐藏的root窗口以满足tkinter初始化
    root = tk.Tk()
    root.withdraw()

    results = []
    results.append(test_match_logic())
    results.append(test_excel_reading())

    root.destroy()

    all_passed = all(results)
    print(f"\n总体结果: {'全部通过 ✅' if all_passed else '有失败项 ❌'}")
    sys.exit(0 if all_passed else 1)
