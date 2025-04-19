#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
示例测试脚本，演示如何使用验收程序
"""

import os
import sys
import subprocess
import argparse
import glob

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="运行进出口文件验收程序")
    parser.add_argument("--packing-list", help="指定采购装箱单文件路径")
    parser.add_argument("--policy-file", help="指定政策文件路径")
    parser.add_argument("--output-dir", help="指定输出目录")
    parser.add_argument("--template-dir", help="指定模板目录")
    parser.add_argument("--report-path", help="指定报告输出路径")
    parser.add_argument("--skip-processing", action="store_true", help="跳过文件处理，仅验证已生成的文件")
    parser.add_argument("--debug", action="store_true", help="启用调试模式，打印更多调试信息")
    args = parser.parse_args()
    
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 定义默认测试文件路径
    default_packing_list = os.path.join(os.path.dirname(current_dir), "testfiles", "original_packing_list.xlsx")
    default_policy_file = os.path.join(os.path.dirname(current_dir), "testfiles", "policy.xlsx")
    default_output_dir = os.path.join(os.path.dirname(current_dir), "outputs")
    default_template_dir = os.path.join(os.path.dirname(current_dir), "Template")
    default_report_path = os.path.join(current_dir, "reports", "validation_report.md")
    
    # 使用命令行参数或默认值
    packing_list_path = args.packing_list or default_packing_list
    policy_file_path = args.policy_file or default_policy_file
    output_dir = args.output_dir or default_output_dir
    template_dir = args.template_dir or default_template_dir
    report_path = args.report_path or default_report_path
    
    # 检查文件是否存在
    if not os.path.exists(packing_list_path):
        print(f"错误: 采购装箱单文件不存在: {packing_list_path}")
        print("请使用 --packing-list 指定正确的文件路径")
        # 查找可能的装箱单文件
        possible_files = []
        search_paths = [".", os.path.dirname(current_dir), os.path.join(os.path.dirname(current_dir), "testfiles")]
        for path in search_paths:
            possible_files.extend(glob.glob(os.path.join(path, "*packing*.xlsx")))
            possible_files.extend(glob.glob(os.path.join(path, "*装箱*.xlsx")))
        
        if possible_files:
            print("\n可能的装箱单文件:")
            for file in possible_files:
                print(f"  - {os.path.abspath(file)}")
        
        sys.exit(1)
        
    if not os.path.exists(policy_file_path):
        print(f"错误: 政策文件不存在: {policy_file_path}")
        print("请使用 --policy-file 指定正确的文件路径")
        # 查找可能的政策文件
        possible_files = []
        search_paths = [".", os.path.dirname(current_dir), os.path.join(os.path.dirname(current_dir), "testfiles")]
        for path in search_paths:
            possible_files.extend(glob.glob(os.path.join(path, "*policy*.xlsx")))
            possible_files.extend(glob.glob(os.path.join(path, "*政策*.xlsx")))
        
        if possible_files:
            print("\n可能的政策文件:")
            for file in possible_files:
                print(f"  - {os.path.abspath(file)}")
                
        sys.exit(1)
        
    if not os.path.exists(template_dir):
        print(f"警告: 模板目录不存在: {template_dir}")
        
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    
    # 构建命令
    cmd = [
        sys.executable,  # 当前Python解释器
        os.path.join(current_dir, "main.py"),
        "--packing-list", packing_list_path,
        "--policy-file", policy_file_path,
        "--output-dir", output_dir,
        "--template-dir", template_dir,
        "--report-path", report_path
    ]
    
    if args.skip_processing:
        cmd.append("--skip-processing")
    
    if args.debug:
        cmd.append("--debug")
    
    # 显示执行的命令
    print("执行命令:", " ".join(cmd))
    
    # 执行命令
    print("\n启动验收程序...")
    try:
        subprocess.run(cmd, check=True)
        print("验收程序执行完成。")
        
        # 检查报告是否生成
        if os.path.exists(report_path):
            print(f"报告已生成: {report_path}")
            
            # 可选: 打开报告
            if sys.platform.startswith('win'):
                os.startfile(report_path)
            elif sys.platform.startswith('darwin'):  # macOS
                subprocess.run(['open', report_path])
            else:  # Linux
                subprocess.run(['xdg-open', report_path])
        else:
            print(f"错误: 报告未生成: {report_path}")
            
    except subprocess.CalledProcessError as e:
        print(f"错误: 验收程序执行失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 