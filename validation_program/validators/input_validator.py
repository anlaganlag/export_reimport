import pandas as pd
import re
import json
import os
from .utils import find_column_with_pattern, read_excel_to_df


class InputValidator:
    """输入文件验证器"""
    
    def __init__(self, config_path=None):
        """初始化验证器"""
        if config_path is None:
            # 默认配置路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(os.path.dirname(current_dir), "config", "validation_rules.json")
        
        # 如果配置文件不存在，使用默认配置
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                self.rules = json.load(f)
        else:
            self.rules = {
                "price_validation": {
                    "decimal_places": {
                        "exchange_rate": 4
                    }
                }
            }
            
        # 设置可跳过的表头行数
        self.skiprows = 0

    def detect_file_structure(self, file_path):
        """检测文件结构并设置跳过的行数
        
        Args:
            file_path: Excel文件路径
            
        Returns:
            int: 应跳过的表头行数
        """
        try:
            # 读取前几行进行检测
            header_rows = pd.read_excel(file_path, nrows=3, header=None)
            
            # 检查第一行是否是文件信息行(例如"装货清单 2025年更新版本")
            first_row = str(header_rows.iloc[0, 0]) if not pd.isna(header_rows.iloc[0, 0]) else ""
            
            # if ("年" in first_row or "版本" in first_row) and len(first_row) < 30:
            if ("采购装箱单" in first_row or "billionaire" in first_row) and len(first_row) < 30:
                self.skiprows = 1
                return 1
                
            # 默认不跳过表头
            self.skiprows = 0
            return 0
        except Exception:
            # 默认值
            self.skiprows = 0
            return 0

    def validate_packing_list_header(self, file_path):
        """验证采购装箱单表头
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 直接读取第一行进行验证，不跳过任何行
            header_rows = pd.read_excel(file_path, nrows=1, header=None)
            # 读取Excel文件第一行第一列(A1单元格)的内容
            header_text = str(header_rows.iloc[0, 0])
            
            # 检查是否包含必要标题文本
            required_headers = ["采购装箱单"]
            found_headers = []
            
            for header in required_headers:
                if header.lower() in header_text.lower():
                    found_headers.append(header)
            
            if not found_headers:
                return {
                    "success": False, 
                    "message": f"表头未包含任何所需标题文本。需要: {', '.join(required_headers)}，实际值: '{header_text[:50]}...'。验收标准: 文件第1行应包含装箱单或相关标题文本(参见文档第12行要求)。"
                }
                
            # 检查是否包含编号
            id_match = re.search(r'[A-Za-z0-9-]+', header_text)
            if not id_match:
                return {
                    "success": False, 
                    "message": f"表头未包含编号。实际值: '{header_text[:50]}...'。验收标准: 文件表头应包含采购单编号(如PL-20250418-0001格式)。正确示例: '采购装箱单 PL-20250418-0001'"
                }
                
            # 检测文件结构（在验证完表头后）
            self.detect_file_structure(file_path)
                
            return {
                "success": True, 
                "message": f"采购装箱单表头验证通过。找到标题: '{', '.join(found_headers)}', 编号: '{id_match.group(0)}'"
            }
        except Exception as e:
            return {"success": False, "message": f"验证表头时出错: {str(e)}。文件路径: {file_path}"}
    
    def extract_id(self, file_path):
        """从采购装箱单提取编号
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            str: 提取的编号，若未找到则返回None
        """
        try:
            header_rows = pd.read_excel(file_path, nrows=1, header=None)
            header_text = str(header_rows.iloc[0, 0])
            match = re.search(r'(\w+-\d+|\w+\d+)', header_text)
            if match:
                return match.group(1)
            return None
        except Exception:
            return None
            
    def validate_packing_list_field_headers(self, file_path):
        """验证表头字段名分两行（中英文）
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 检测文件结构如果还没检测过
            if self.skiprows == 0:
                self.detect_file_structure(file_path)
                
            # 读取表头行
            start_row = self.skiprows
            header_df = pd.read_excel(file_path, header=None, skiprows=start_row, nrows=4)
            
            # 通常字段名在第1,2行(经过skiprows处理后)
            english_row = header_df.iloc[0]  # 跳过行后的第一行
            chinese_row = header_df.iloc[1]  # 跳过行后的第二行
            
            # 必须字段列表
            required_chinese =  ["序号", "料号", "供应商", "项目名称", "工厂地点", "进口清关货描", "供应商开票名称", "物料名称", "型号", "数量", "单位", "纸箱尺寸", "单件体积", "总体积", "单件毛重", "总毛重", "单件净重", "总净重", "每箱数量", "总件数", "箱号", "栈板尺寸", "栈板编号", "出口报关方式", "采购公司", "采购单价不含税", "开票税率"]

            required_english =  ["S/N", "Part Number", "Supplier", "Project", "Plant Location", "Commodity Description (Customs)", "Commercial Invoice Description", "EPR Part NameEPR", "Model Number", "Quantity", "Unit", "Carton Size (L×W×H in mm)", "Unit Volume (CBM)", "Total Volume (CBM)", "Gross Weight per Unit (kg)", "Total Gross Weight (kg)", "Net Weight per Unit (kg)", "Total Net Weight (kg)", "Quantity per Carton", "Total Carton Quantity", "Carton Number", "Pallet Size (L×W×H in mm)", "Pallet ID", "Export Declaration Method", "Purchasing Company", "Unit Price (Excl. Tax, CNY)()", "Tax Rate (%)"]
            
            # 如果发现表头结构不同，尝试其他组合
            if not any(req.lower() in str(x).lower() for x in english_row if pd.notna(x) for req in required_english):
                english_row = header_df.iloc[1]  # 尝试第二行
                chinese_row = header_df.iloc[2]  # 尝试第三行
            
            # 检查是否有足够的非空字段
            chinese_fields = chinese_row[chinese_row.notna()].count()
            english_fields = english_row[english_row.notna()].count()
            
            if chinese_fields < 5 or english_fields < 5:
                return {
                    "success": False, 
                    "message": f"表头字段行不完整。中文字段数: {chinese_fields}, 英文字段数: {english_fields}, 需要至少5个。验收标准: 表头应包含中英文字段名(参见文档第13行要求)。"
                }
                
            # 检查是否有必须的中文字段
            chinese_fields_found = []
            missing_chinese = []
            
            for field in chinese_row[chinese_row.notna()]:
                field_str = str(field).strip()
                chinese_fields_found.append(field_str)
                
            for req in required_chinese:
                if not any(req in field for field in chinese_fields_found):
                    missing_chinese.append(req)
            
            # 检查是否有必须的英文字段
            english_fields_found = []
            missing_english = []
            
            for field in english_row[english_row.notna()]:
                field_str = str(field).strip()
                english_fields_found.append(field_str)
                
            for req in required_english:
                if not any(req.lower() in field.lower() for field in english_fields_found):
                    missing_english.append(req)
            
            # 验证结果
            if missing_chinese:
                return {
                    "success": False, 
                    "message": f"缺少必要的中文字段: {', '.join(missing_chinese)}。找到的字段: {', '.join(chinese_fields_found[:5])}...。验收标准: 表头应包含所有必要的中文字段(参见文档第14行要求)。"
                }
                
            if missing_english:
                return {
                    "success": False, 
                    "message": f"缺少必要的英文字段: {', '.join(missing_english)}。找到的字段: {', '.join(english_fields_found[:5])}...。验收标准: 表头应包含所有必要的英文字段(参见文档第14行要求)。"
                }
                
            return {
                "success": True, 
                "message": f"表头字段名验证通过。中文字段: {len(chinese_fields_found)}个, 英文字段: {len(english_fields_found)}个"
            }
        except Exception as e:
            return {"success": False, "message": f"验证字段头时出错: {str(e)}。文件路径: {file_path}"}
    
    def validate_weights(self, file_path):
        """验证每行数据的单件总净重小于总毛重
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 跳过表头，读取数据行
            df = pd.read_excel(file_path, skiprows=1)
            
            # 查找净重和毛重列
            net_weight_col = find_column_with_pattern(df, ["Total Gross Weight (kg)"])
            gross_weight_col = find_column_with_pattern(df, ["Total Gross Weight (kg)"])
            
            if net_weight_col is None or gross_weight_col is None:
                return {"success": False, "message": "未找到净重或毛重列"}
            
            # 检查每行净重是否小于毛重
            invalid_rows = []
            for idx, row in df.iterrows():
                # 跳过汇总行和空行
                if pd.isna(row[net_weight_col]) or pd.isna(row[gross_weight_col]):
                    continue
                    
                if row[net_weight_col] > row[gross_weight_col]:
                    invalid_rows.append(idx + 4)  # +4是因为跳过了3行表头，再加上1行零基索引
            
            if invalid_rows:
                return {
                    "success": False, 
                    "message": f"以下行的净重大于毛重: {', '.join(map(str, invalid_rows))}"
                }
                
            return {"success": True, "message": "净重毛重验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证净重毛重时出错: {str(e)}"}
    
    def validate_summary_data(self, file_path):
        """验证表尾汇总数据正确性
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取所有数据
            df = pd.read_excel(file_path, skiprows=3)
            
            # 查找数量、体积、毛重、净重列
            qty_col = find_column_with_pattern(df, ["数量", "Quantity"])
            vol_col = find_column_with_pattern(df, ["体积", "Volume"])
            gross_col = find_column_with_pattern(df, ["毛重", "Gross Weight"])
            net_col = find_column_with_pattern(df, ["净重", "Net Weight"])
            
            if not all([qty_col, vol_col, gross_col, net_col]):
                return {"success": False, "message": "未找到所有需要验证的列"}
            
            # 找到汇总行（通常表尾最后几行）
            summary_row = None
            for i in range(len(df) - 1, max(0, len(df) - 10), -1):
                # 汇总行通常有"合计"或"总计"字样，或者第一列为空而数量列有值
                if (isinstance(df.iloc[i, 0], str) and 
                    ("合计" in df.iloc[i, 0] or "总计" in df.iloc[i, 0] or "Total" in df.iloc[i, 0])):
                    summary_row = i
                    break
            
            if summary_row is None:
                return {"success": False, "message": "未找到汇总行"}
            
            # 获取汇总值
            summary_qty = df.iloc[summary_row][qty_col]
            summary_vol = df.iloc[summary_row][vol_col]
            summary_gross = df.iloc[summary_row][gross_col]
            summary_net = df.iloc[summary_row][net_col]
            
            # 计算实际值
            # 排除汇总行和空行
            data_rows = df.iloc[:summary_row].copy()
            data_rows = data_rows[data_rows[qty_col].notna()]
            
            actual_qty = data_rows[qty_col].sum()
            actual_vol = data_rows[vol_col].sum() if vol_col and vol_col in data_rows else 0
            actual_gross = data_rows[gross_col].sum() if gross_col and gross_col in data_rows else 0
            actual_net = data_rows[net_col].sum() if net_col and net_col in data_rows else 0
            
            # 比较汇总值与实际计算值
            precision = 0.01  # 允许的误差
            errors = []
            
            if abs(summary_qty - actual_qty) > precision:
                errors.append(f"数量汇总不正确: 显示{summary_qty}, 实际{actual_qty}")
                
            if abs(summary_vol - actual_vol) > precision:
                errors.append(f"体积汇总不正确: 显示{summary_vol}, 实际{actual_vol}")
                
            if abs(summary_gross - actual_gross) > precision:
                errors.append(f"毛重汇总不正确: 显示{summary_gross}, 实际{actual_gross}")
                
            if abs(summary_net - actual_net) > precision:
                errors.append(f"净重汇总不正确: 显示{summary_net}, 实际{actual_net}")
            
            if errors:
                return {"success": False, "message": "汇总数据有误: " + "; ".join(errors)}
                
            return {"success": True, "message": "汇总数据验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证汇总数据时出错: {str(e)}"}
    
    def validate_policy_file_id(self, policy_file_path, packing_list_id):
        """验证政策文件编号与采购装箱单编号一致
        
        Args:
            policy_file_path: 政策文件路径
            packing_list_id: 采购装箱单编号
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            if packing_list_id is None:
                return {
                    "success": False, 
                    "message": "采购装箱单编号未提取到。验收标准: 采购装箱单必须包含可识别的编号。"
                }
                
            # 读取政策文件
            policy_df = pd.read_excel(policy_file_path)
            
            # 查找编号列或表头
            id_found = False
            policy_id = None
            
            # 先检查表头
            header_rows = pd.read_excel(policy_file_path, nrows=1, header=None)
            header_text = str(header_rows.iloc[0, 0])
            match = re.search(r'(\w+-\d+|\w+\d+)', header_text)
            if match:
                policy_id = match.group(1)
                id_found = True
            
            # 如果表头没找到，检查编号列
            if not id_found:
                id_cols = ["编号", "Number", "ID", "单号", "装箱单号", "Policy No", "参考编号"]
                for col in id_cols:
                    if col in policy_df.columns and not policy_df[col].empty:
                        policy_id = str(policy_df[col].iloc[0])
                        id_found = True
                        break
                        
                # 尝试查找包含这些关键词的列
                if not id_found:
                    for col in policy_df.columns:
                        if any(key in str(col) for key in ["编号", "号", "ID", "No"]):
                            policy_id = str(policy_df[col].iloc[0])
                            id_found = True
                            break
            
            if not id_found or policy_id is None:
                return {
                    "success": False, 
                    "message": f"未在政策文件中找到编号。验收标准: 政策文件应包含与装箱单匹配的编号(参见文档第21行要求)。正确示例: 表头包含'Policy No: PL-20250418-0001'或有单独的'编号'列。"
                }
                
            # 比较编号是否一致
            if packing_list_id not in policy_id and policy_id not in packing_list_id:
                return {
                    "success": False, 
                    "message": f"政策文件编号({policy_id})与采购装箱单编号({packing_list_id})不一致。验收标准: 两个文件的编号应保持一致(参见文档第21行要求)。"
                }
                
            return {"success": True, "message": "政策文件编号验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证政策文件编号时出错: {str(e)}。文件路径: {policy_file_path}"}
    
    def validate_exchange_rate_decimal(self, policy_file_path):
        """验证汇率保留4位小数
        
        Args:
            policy_file_path: 政策文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取政策文件
            policy_df = pd.read_excel(policy_file_path)
            
            # 查找汇率列
            exchange_col = find_column_with_pattern(policy_df, ["汇率", "Exchange Rate"])
            
            if exchange_col is None:
                return {"success": False, "message": "未找到汇率列"}
            
            # 获取汇率值
            exchange_rates = policy_df[exchange_col].dropna()
            if exchange_rates.empty:
                return {"success": False, "message": "汇率列无数据"}
            
            # 验证是否保留4位小数
            invalid_rates = []
            for rate in exchange_rates:
                # 将数字转为字符串检查小数位数
                rate_str = str(rate)
                if '.' in rate_str:
                    decimal_part = rate_str.split('.')[1]
                    if len(decimal_part) != 4:
                        invalid_rates.append(rate)
            
            if invalid_rates:
                return {
                    "success": False, 
                    "message": f"以下汇率未保留4位小数: {', '.join(map(str, invalid_rates))}"
                }
                
            return {"success": True, "message": "汇率小数位验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证汇率小数位时出错: {str(e)}"}
    
    def validate_company_bank_info(self, policy_file_path):
        """验证政策文件包含公司信息和银行信息
        
        Args:
            policy_file_path: 政策文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取政策文件为DataFrame
            policy_df = pd.read_excel(policy_file_path)
            
            # 将DataFrame转换为字符串以便于搜索
            # 合并所有单元格内容为一个大字符串
            all_text = ""
            
            # 检查所有列名
            for col in policy_df.columns:
                all_text += str(col) + " "
            
            # 检查所有单元格内容
            for _, row in policy_df.iterrows():
                for item in row:
                    all_text += str(item) + " "
            
            # 检查公司信息
            company_patterns = ["公司", "Company", "企业", "Enterprise"]
            has_company_info = any(pattern in all_text for pattern in company_patterns)
            
            # 检查银行信息
            bank_patterns = ["银行", "Bank", "账号", "Account"]
            has_bank_info = any(pattern in all_text for pattern in bank_patterns)
            
            if not has_company_info:
                return {"success": False, "message": "政策文件未包含公司信息"}
                
            if not has_bank_info:
                return {"success": False, "message": "政策文件未包含银行信息"}
                
            return {"success": True, "message": "公司信息和银行信息验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证公司银行信息时出错: {str(e)}"}
    
    def validate_all(self, packing_list_path, policy_file_path):
        """运行所有输入验证
        
        Args:
            packing_list_path: 采购装箱单文件路径
            policy_file_path: 政策文件路径
            
        Returns:
            dict: 包含所有验证结果的字典
        """
        results = {}
        
        # 采购装箱单验证
        results["packing_list_header"] = self.validate_packing_list_header(packing_list_path)
        results["packing_list_field_headers"] = self.validate_packing_list_field_headers(packing_list_path)
        results["weights"] = self.validate_weights(packing_list_path)
        results["summary_data"] = self.validate_summary_data(packing_list_path)
        
        # 提取采购装箱单编号
        packing_list_id = self.extract_id(packing_list_path)
        
        # 政策文件验证
        results["policy_file_id"] = self.validate_policy_file_id(policy_file_path, packing_list_id)
        results["exchange_rate_decimal"] = self.validate_exchange_rate_decimal(policy_file_path)
        results["company_bank_info"] = self.validate_company_bank_info(policy_file_path)
        
        return results 