#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文选题查重系统 - 主程序
"""

import sys
import os
from pathlib import Path
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from checker import ThesisChecker


def main():
    """主函数"""
    try:
        print("=" * 70)
        print("论文选题查重系统 v2.0")
        print("=" * 70)
        print()

        # 检查数据文件夹
        data_dir = Path("原始数据")
        if not data_dir.exists():
            print("\n正在创建必要的文件夹...")
            data_dir.mkdir(parents=True, exist_ok=True)
            print("✓ 已创建「原始数据」文件夹")
            print("✓ 已创建「二次提交」文件夹")
            print("✓ 已创建「查重结果」文件夹")
            print()
            print("请将待检测的文件放入相应文件夹后，再运行程序")
            return

        # 创建查重器实例
        checker = ThesisChecker(
            data_dir="原始数据",
            resubmit_dir="二次提交",
            output_dir="查重结果",
            threshold=0.75
        )

        # 运行查重
        checker.run()

        print()
        print("=" * 70)
        print("查重完成！")
        print("=" * 70)

    except FileNotFoundError as e:
        print(f"错误：{e}")
        print("请确保数据文件夹存在且包含文件")
        sys.exit(1)
    except Exception as e:
        print(f"程序运行出错：{e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
