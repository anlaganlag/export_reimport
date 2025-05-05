#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
测试验证程序
"""

import os
import subprocess
import sys

def main():
    """主函数"""
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 定义测试文件路径
    packing_list_path = os.path.join(current_dir, "testfiles", "original_packing_list.xlsx")
    policy_file_path = os.path.join(current_dir, "testfiles", "policy.xlsx")
    output_dir = os.path.join(current_dir, "outputs")
    
    # 构建命令
    cmd = [
        sys.executable,  # 当前Python解释器
        os.path.join(current_dir, "validation_program", "run_validation.py"),
        "--packing-list", packing_list_path,
        "--policy-file", policy_file_path,
        "--output-dir", output_dir,
        "--debug"
    ]
    
    # 显示执行的命令
    print("执行命令:", " ".join(cmd))
    
    # 执行命令
    print("\n启动验收程序...")
    try:
        subprocess.run(cmd, check=True)
        print("验收程序执行完成。")
    except subprocess.CalledProcessError as e:
        print(f"错误: 验收程序执行失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
