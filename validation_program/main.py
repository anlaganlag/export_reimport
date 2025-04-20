#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import argparse
import pandas as pd
import json
import sys
import glob
from validators.input_validator import InputValidator
from validators.output_validator import OutputValidator
from validators.process_validator import ProcessValidator
from validators.utils import get_output_files

# 导入process_shipping_list模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import process_shipping_list


def generate_report(results, report_path, args):
    """生成验收报告
    
    Args:
        results: 包含验证结果的字典
        report_path: 报告输出路径
        args: 命令行参数
    """
    # 确保输出目录存在
    os.makedirs(os.path.dirname(report_path), exist_ok=True)
    
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# 进出口文件生成系统验收报告\n\n")
        
        # 添加验证说明
        f.write("## 验证说明\n\n")
        f.write("本报告包含对进出口文件生成系统的完整验证结果。验证过程包括以下几个部分：\n\n")
        f.write("1. **输入文件验证** - 检查装箱单和政策文件的格式和内容是否符合要求\n")
        f.write("2. **处理逻辑验证** - 验证文件处理逻辑是否正确\n")
        f.write("3. **输出文件验证** - 验证生成的输出文件是否符合要求\n\n")
        f.write("如果任何验证项目失败，整体验收结果将被标记为\"不通过\"。请查看详细测试结果部分以找出具体问题。\n\n")
        
        # 文件信息部分
        f.write("## 验证文件信息\n\n")
        f.write("### 输入文件\n")
        f.write(f"- 采购装箱单: `{os.path.abspath(args.packing_list)}`\n")
        f.write(f"- 政策文件: `{os.path.abspath(args.policy_file)}`\n\n")
        
        f.write("### 输出目录\n")
        f.write(f"- 输出路径: `{os.path.abspath(args.output_dir)}`\n")
        if args.template_dir:
            f.write(f"- 模板目录: `{os.path.abspath(args.template_dir)}`\n")
        f.write(f"- 报告路径: `{os.path.abspath(report_path)}`\n\n")
        
        # 处理参数
        f.write("### 处理参数\n")
        f.write(f"- 跳过文件处理: {'是' if args.skip_processing else '否'}\n\n")
        
        # 检测输出文件
        output_files = []
        if os.path.exists(args.output_dir):
            output_files = [f for f in os.listdir(args.output_dir) if f.endswith(('.xlsx', '.csv', '.pdf'))]
        
        if output_files:
            f.write("### 检测到的输出文件\n")
            for file in sorted(output_files):
                f.write(f"- `{file}`\n")
            f.write("\n")
        
        # 验收标准部分
        f.write("## 验收标准\n\n")
        f.write("### 输入文件标准\n\n")
        f.write("1. **采购装箱单标题**\n")
        f.write("   - 要求: 文件应包含'采购装箱单'或'装箱单'等相关标题文本\n")
        f.write("   - 格式示例: '采购装箱单 PL-20250418-0001'\n")
        f.write("   - 参考文档: testfiles/README.md 第12行\n\n")
        
        f.write("2. **字段头格式**\n")
        f.write("   - 要求: 表头应包含中英文字段名，含必要字段\n")
        f.write("   - 必要中文字段: 序号、零件号、描述、数量、单位、净重、毛重\n")
        f.write("   - 必要英文字段: No、Part、Description、Quantity、Unit、Net、Gross\n")
        f.write("   - 参考文档: testfiles/README.md 第13-14行\n\n")
        
        f.write("3. **政策文件要求**\n")
        f.write("   - 要求: 政策文件应包含与装箱单匹配的编号，以及完整的参数设置\n")
        f.write("   - 必要内容: 匹配编号、汇率、加价率、保险费率、公司和银行信息\n")
        f.write("   - 参考文档: testfiles/README.md 第21-25行\n\n")
        
        # 整体验收结果
        all_passed = all(result["success"] for result in results.values())
        f.write(f"## 验收结果：{'通过' if all_passed else '不通过'}\n\n")
        
        # 按类别统计通过/失败数量
        categories = {
            "输入文件验证": [k for k in results.keys() if k.startswith(("packing_list", "policy", "weights", "summary_data", "exchange_rate_decimal", "company_bank_info"))],
            "处理逻辑验证": [k for k in results.keys() if any(k.startswith(p) for p in ["trade_type", "price", "merge", "split", "fob", "insurance", "freight", "cif"])],
            "输出文件验证": [k for k in results.keys() if any(k.startswith(p) for p in ["export", "import", "file", "format"])]
        }
        
        # 将任何未分类的验证结果添加到适当的类别中
        all_result_keys = set(results.keys())
        categorized_keys = set()
        for keys in categories.values():
            categorized_keys.update(keys)
        
        uncategorized_keys = all_result_keys - categorized_keys
        if uncategorized_keys:
            # 尝试将未分类的键添加到合适的类别中
            for key in uncategorized_keys:
                if "file" in key or "packing" in key or "policy" in key:
                    categories["输入文件验证"].append(key)
                elif "process" in key:
                    categories["处理逻辑验证"].append(key)
                else:
                    # 如果无法确定类别，添加到输入文件验证
                    categories["输入文件验证"].append(key)
        
        f.write("## 验证结果统计\n\n")
        for category, keys in categories.items():
            if not keys:
                continue
                
            category_results = {k: results[k] for k in keys if k in results}
            passed = sum(1 for v in category_results.values() if v["success"])
            total = len(category_results)
            
            f.write(f"### {category}: {passed}/{total} 通过\n\n")
        
        # 详细测试结果
        f.write("## 详细测试结果\n\n")
        
        # 按类别分组显示结果
        for category, keys in categories.items():
            if not keys:
                continue
                
            f.write(f"### {category}\n\n")
            
            # 先显示失败的测试
            failed_tests = [key for key in sorted(keys) if key in results and not results[key]["success"]]
            passed_tests = [key for key in sorted(keys) if key in results and results[key]["success"]]
            
            if failed_tests:
                f.write("#### ❌ 失败的测试\n\n")
                for key in failed_tests:
                    result = results[key]
                    f.write(f"##### {key}: ❌ 失败\n")
                    f.write(f"- 结果: {result['message']}\n\n")
            
            if passed_tests:
                f.write("#### ✅ 通过的测试\n\n")
                for key in passed_tests:
                    result = results[key]
                    f.write(f"##### {key}: ✅ 通过\n")
                    f.write(f"- 结果: {result['message']}\n\n")
            
            # 显示有关本类别测试的统计信息
            category_results = {k: results[k] for k in keys if k in results}
            passed = sum(1 for v in category_results.values() if v["success"])
            total = len(category_results)
            f.write(f"**统计**: {passed}/{total} 测试通过\n\n")
        
        # 添加解决方案建议
        if not all_passed:
            f.write("## 解决方案建议\n\n")
            
            for category, keys in categories.items():
                category_failures = [(k, results[k]) for k in keys if k in results and not results[k]["success"]]
                if category_failures:
                    f.write(f"### {category}问题修复建议\n\n")
                    
                    for key, result in category_failures:
                        f.write(f"#### {key}:\n")
                        
                        # 针对不同问题提供具体建议
                        if "表头未包含任何所需标题文本" in result["message"]:
                            f.write("- **建议**: 修改文件首行，确保包含'采购装箱单'、'装箱单'或'PACKING LIST'等关键词\n")
                            f.write("- **正确示例**: '采购装箱单 PL-20250418-0001'\n")
                        elif "表头未包含编号" in result["message"]:
                            f.write("- **建议**: 在文件首行添加编号，通常跟在标题后面\n")
                            f.write("- **正确示例**: '采购装箱单 PL-20250418-0001'\n")
                        elif "缺少必要的中文字段" in result["message"] or "缺少必要的英文字段" in result["message"]:
                            f.write("- **建议**: 检查文件表头行，确保包含所有必要的中英文字段\n")
                            f.write("- **中文字段**: 序号、零件号、描述、数量、单位、净重、毛重\n")
                            f.write("- **英文字段**: No、Part No.、Description、Quantity、Unit、Net Weight、Gross Weight\n")
                        elif "未在政策文件中找到编号" in result["message"]:
                            f.write("- **建议**: 在政策文件中添加与装箱单匹配的编号\n")
                            f.write("- **正确示例**: 添加标题行'Policy No: PL-20250418-0001'或添加'编号'列\n")
                        elif "政策文件编号与采购装箱单编号不一致" in result["message"]:
                            f.write("- **建议**: 确保政策文件中的编号与采购装箱单编号一致\n")
                        elif "验证表头时出错" in result["message"] or "验证字段头时出错" in result["message"]:
                            f.write("- **建议**: 检查文件格式是否正确，特别是表头部分的格式\n")
                            f.write("- **注意**: 如果第一行是文件信息或版本信息，应确保实际表头内容从第二行开始\n")
                        elif "净重大于毛重" in result["message"] or "以下行的净重大于毛重" in result["message"]:
                            f.write("- **建议**: 检查指定行的毛重和净重值，确保每行的净重小于或等于毛重\n")
                            f.write("- **原因**: 根据物理原则，毛重（含包装重量）应大于等于净重（不含包装）\n")
                        elif "汇总数据有误" in result["message"]:
                            f.write("- **建议**: 检查表底汇总行的计算，确保数量、体积、净重和毛重的总和正确\n")
                            f.write("- **解决方法**: 重新计算各列的总和，并更新汇总行的值\n")
                        elif "汇率未保留4位小数" in result["message"] or "以下汇率未保留4位小数" in result["message"]:
                            f.write("- **建议**: 修改政策文件中的汇率值，确保所有汇率保留4位小数\n")
                            f.write("- **正确示例**: 6.8976 (四位小数)\n")
                        elif "未找到汇率列" in result["message"] or "汇率列无数据" in result["message"]:
                            f.write("- **建议**: 检查政策文件中是否包含汇率列，并确保有正确的数据\n")
                            f.write("- **列名示例**: '汇率' 或 'Exchange Rate'\n")
                        elif "政策文件未包含公司信息" in result["message"]:
                            f.write("- **建议**: 在政策文件中添加必要的公司信息\n")
                            f.write("- **必要信息**: 公司名称、地址、联系方式等\n")
                        elif "政策文件未包含银行信息" in result["message"]:
                            f.write("- **建议**: 在政策文件中添加必要的银行信息\n")
                            f.write("- **必要信息**: 银行名称、账号、Swift代码等\n")
                        elif "验证汇总数据时出错" in result["message"]:
                            f.write("- **建议**: 检查装箱单格式，确保汇总行正确且可被识别\n")
                            f.write("- **注意**: 汇总行通常在表格底部，包含'合计'或'总计'字样\n")
                        else:
                            f.write(f"- **建议**: 根据错误信息修复问题: {result['message']}\n")
                        
                        f.write("\n")
        
        # 添加验证时间戳
        import datetime
        f.write(f"\n\n---\n生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


def find_cif_invoice(output_dir):
    """查找CIF原始发票文件
    
    Args:
        output_dir: 输出目录路径
        
    Returns:
        CIF发票文件路径或None
    """
    # 尝试不同可能的命名
    patterns = [
        os.path.join(output_dir, "cif_original_invoice.xlsx"),
        os.path.join(output_dir, "cif_invoice.xlsx"),
        os.path.join(output_dir, "*cif*.xlsx")
    ]
    
    for pattern in patterns:
        files = glob.glob(pattern)
        if files:
            return files[0]
    
    return None


def find_export_invoice(output_dir):
    """查找出口发票文件
    
    Args:
        output_dir: 输出目录路径
        
    Returns:
        出口发票文件路径或None
    """
    # 尝试不同可能的命名
    patterns = [
        os.path.join(output_dir, "出口-*.xlsx"),
        os.path.join(output_dir, "export_invoice.xlsx"),
        os.path.join(output_dir, "export*.xlsx")
    ]
    
    for pattern in patterns:
        files = glob.glob(pattern)
        if files:
            return files[0]
    
    return None


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="进出口文件生成系统验收程序")
    parser.add_argument("--packing-list", required=True, help="采购装箱单路径")
    parser.add_argument("--policy-file", required=True, help="政策文件路径")
    parser.add_argument("--output-dir", default="outputs", help="输出目录")
    parser.add_argument("--template-dir", default="Template", help="模板目录")
    parser.add_argument("--report-path", default="reports/validation_report.md", help="报告输出路径")
    parser.add_argument("--skip-processing", action="store_true", help="跳过文件处理，仅验证已生成的文件")
    
    args = parser.parse_args()
    
    # 确保输出目录存在
    os.makedirs(args.output_dir, exist_ok=True)
    os.makedirs(os.path.dirname(args.report_path), exist_ok=True)
    
    # 初始化验证器
    input_validator = InputValidator()
    process_validator = ProcessValidator()
    output_validator = OutputValidator()
    
    # 结果收集
    results = {}
    
    # 1. 验证输入文件
    print("正在验证输入文件...")
    input_results = input_validator.validate_all(args.packing_list, args.policy_file)
    results.update(input_results)
    
    # 检查输入验证是否通过
    input_valid = all(v["success"] for k, v in input_results.items())
    if not input_valid:
        print("输入文件验证失败，生成报告并退出...")
        generate_report(results, args.report_path, args)
        return
    
    # 2. 调用处理程序
    if not args.skip_processing:
        print("正在处理文件...")
        try:
            process_shipping_list.process_shipping_list(
                args.packing_list,
                args.policy_file,
                args.output_dir
            )
            print("文件处理完成。")
        except Exception as e:
            results["process_error"] = {"success": False, "message": f"处理文件时出错: {str(e)}"}
            generate_report(results, args.report_path, args)
            print(f"错误: {str(e)}")
            return
    
    # 3. 查找关键文件
    print("正在查找输出文件...")
    cif_invoice_path = find_cif_invoice(args.output_dir)
    export_invoice_path = find_export_invoice(args.output_dir)
    
    # 4. 验证处理逻辑
    print("正在验证处理逻辑...")
    if cif_invoice_path:
        process_results = process_validator.validate_all(
            args.packing_list,
            args.policy_file,
            cif_invoice_path,
            export_invoice_path,
            args.output_dir
        )
        results.update(process_results)
    else:
        results["process_missing_cif"] = {"success": False, "message": "未找到CIF原始发票文件，无法验证处理逻辑"}
    
    # 5. 验证输出文件
    print("正在验证输出文件...")
    output_results = output_validator.validate_all(
        args.output_dir,
        args.packing_list,
        args.template_dir
    )
    results.update(output_results)
    
    # 6. 生成报告
    print(f"正在生成验收报告 {args.report_path}...")
    generate_report(results, args.report_path, args)
    
    # 7. 输出整体结果
    all_passed = all(result["success"] for result in results.values())
    status = "通过" if all_passed else "不通过"
    print(f"验收结果: {status}")
    print(f"详细报告已生成: {args.report_path}")


if __name__ == "__main__":
    main() 