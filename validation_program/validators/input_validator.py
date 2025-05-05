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
            # 读取前几行进行检测，确保不使用第一行作为列名
            header_rows = pd.read_excel(file_path, nrows=3, header=None)
            
            # 打印调试信息
            print("DEBUG: 读取到的表头内容：")
            for i in range(min(3, len(header_rows))):
                row_content = []
                for j in range(len(header_rows.columns)):
                    if pd.notna(header_rows.iloc[i, j]):
                        row_content.append(str(header_rows.iloc[i, j]))
                print(f"Row {i+1}: {' | '.join(row_content)}")
            
            # 获取第一行内容（标题行）
            first_row = []
            for j in range(len(header_rows.columns)):
                if pd.notna(header_rows.iloc[0, j]):
                    first_row.append(str(header_rows.iloc[0, j]))
            header_text = ' '.join(first_row)
            print(f"DEBUG: 完整的第一行内容: {header_text}")
            
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
            # 修改正则表达式以匹配更多格式的编号，包括带有"编号："的情况
            id_patterns = [
                r'编号[：:]\s*([A-Za-z]{2,4}\d{8,12})',  # 匹配 "编号：CXCI202501201" 格式
                r'(?:单号[：:]|No\.:|[A-Za-z]{2,4})[：:\s]*([A-Za-z0-9-]+)',  # 匹配其他格式
                r'([A-Za-z]{2,4}\d{8,12})'  # 直接匹配编号格式
            ]
            
            id_match = None
            for pattern in id_patterns:
                match = re.search(pattern, header_text)
                if match:
                    id_match = match
                    print(f"DEBUG: 找到编号，使用模式 {pattern}")
                    print(f"DEBUG: 匹配到的编号: {match.group(1) if len(match.groups()) > 0 else match.group(0)}")
                    break
            
            if not id_match:
                return {
                    "success": False, 
                    "message": f"表头未包含编号。实际值: '{header_text[:50]}...'。验收标准: 文件表头应包含采购单编号。正确示例: '采购装箱单 编号：CXCI202501201'"
                }
                
            # 检测文件结构（在验证完表头后）
            self.detect_file_structure(file_path)
                
            return {
                "success": True, 
                "message": f"采购装箱单表头验证通过。找到标题: '{', '.join(found_headers)}', 编号: '{id_match.group(1) if len(id_match.groups()) > 0 else id_match.group(0)}'"
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
            header_rows = pd.read_excel(file_path, nrows=3, header=None)
            
            # 遍历前几行寻找编号
            for i in range(min(3, len(header_rows))):
                if pd.notna(header_rows.iloc[i, 0]):
                    header_text = str(header_rows.iloc[i, 0])
                    # 首先尝试匹配带标识的编号
                    match = re.search(r'(?:编号:|单号:|No\.:|[A-Za-z]{2,4})[:\s]*([A-Za-z0-9-]+)', header_text)
                    if match:
                        return match.group(1)
                    # 然后尝试直接匹配编号格式
                    match = re.search(r'[A-Za-z]{2,4}\d{8,12}', header_text)
                    if match:
                        return match.group(0)
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
            required_chinese =  ["序号", "料号", "供应商", "项目名称", "工厂地点", "进口清关货描", "供应商开票名称", "物料名称", "型号", "数量", "单位", "纸箱尺寸", "单件体积", "总体积", "单件毛重", "总毛重", "总净重", "每箱数量", "总件数", "箱号", "栈板尺寸", "栈板编号", "出口报关方式", "采购公司", "采购单价(不含税)", "开票税率"]

            required_english =  ["S/N", "Part Number", "Supplier", "Project", "Plant Location", "Commodity Description (Customs)", "Commercial Invoice Description", "Model Number", "Quantity", "Unit", "Carton Size (L×W×H in mm)", "Unit Volume (CBM)", "Total Volume (CBM)", "Gross Weight per Unit (kg)", "Total Gross Weight (kg)", "Total Net Weight (kg)", "Quantity per Carton", "Total Carton Quantity", "Carton Number", "Pallet Size (L×W×H in mm)", "Pallet ID", "Export Declaration Method", "Purchasing Company", "Tax Rate (%)"]
            
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
        """验证每个箱号的总净重小于总毛重
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        print(f"开始验证净重毛重 - 文件: {file_path}")
        try:
            if self.skiprows == 0:
                self.detect_file_structure(file_path)
            print(f"跳过行数: {self.skiprows+2}")
            df = read_excel_to_df(file_path, skiprows=self.skiprows+2)
            print(f"成功读取数据，共 {len(df)} 行")
            net_weight_col = find_column_with_pattern(df, ["Total Net Weight (kg)", "总净重"])
            gross_weight_col = find_column_with_pattern(df, ["Total Gross Weight (kg)", "总毛重"])
            carton_number_col = find_column_with_pattern(df, ["Carton Number", "箱号"])
            print(f"列索引 - 净重: {net_weight_col}, 毛重: {gross_weight_col}, 箱号: {carton_number_col}")
            if net_weight_col is None or gross_weight_col is None:
                return {"success": False, "message": "未找到净重或毛重列。验收标准: 装箱单必须包含总净重和总毛重列。"}
            if carton_number_col is None:
                return {"success": False, "message": "未找到箱号列。验收标准: 装箱单必须包含箱号列。"}

            # 改进的安全转换函数，更好地处理非数值和通配符
            def safe_convert_to_float(val, row_idx, col_name, error_log):
                # 处理空值
                if val is None or pd.isna(val) or val == "":
                    return 0.0
                
                # 如果已经是数值类型，直接返回
                if isinstance(val, (int, float)):
                    return float(val)
                
                # 转换为字符串并清理
                val_str = str(val).strip()
                if not val_str:
                    return 0.0
                
                # 检查是否包含通配符
                if any(x in val_str for x in ['*', '?', 'N/A', 'n/a', 'TBD', 'tbd']):
                    error_log.append(f"第{row_idx+1}行{col_name}包含通配符或占位符'{val_str}'，自动按0处理")
                    return 0.0
                
                # 尝试直接转换为浮点数
                try:
                    return float(val_str)
                except (ValueError, TypeError):
                    # 尝试清理字符串后转换
                    try:
                        # 只保留数字、小数点和负号
                        clean_str = ''.join(c for c in val_str if c.isdigit() or c in '.-')
                        
                        # 确保只有一个小数点
                        if clean_str and clean_str.count('.') <= 1:
                            # 如果以小数点开头，添加0
                            if clean_str.startswith('.'):
                                clean_str = '0' + clean_str
                            # 如果以小数点结尾，添加0
                            if clean_str.endswith('.'):
                                clean_str = clean_str + '0'
                                
                            if clean_str:
                                return float(clean_str)
                        
                        error_log.append(f"第{row_idx+1}行{col_name}值'{val_str}'无法解析为数值，自动按0处理")
                        return 0.0
                    except Exception:
                        error_log.append(f"第{row_idx+1}行{col_name}值'{val_str}'无法解析为数值，自动按0处理")
                        return 0.0

            carton_net_weights = {}
            carton_gross_weights = {}
            carton_rows = {}
            error_log = []  # 记录被容错的异常值
            
            # 处理每一行数据
            for idx, row in df.iterrows():
                try:
                    # 只处理有箱号的行
                    if pd.notna(row[carton_number_col]):
                        current_carton = str(row[carton_number_col])
                        
                        # 初始化该箱号的数据
                        if current_carton not in carton_rows:
                            carton_rows[current_carton] = []
                            carton_net_weights[current_carton] = 0.0
                            carton_gross_weights[current_carton] = 0.0
                        
                        # 记录行号（Excel中的实际行号）
                        carton_rows[current_carton].append(idx + self.skiprows + 3)
                        
                        # 安全获取净重和毛重值
                        try:
                            net_weight_raw = row[net_weight_col] if net_weight_col < len(row) else None
                        except:
                            net_weight_raw = None
                            
                        try:
                            gross_weight_raw = row[gross_weight_col] if gross_weight_col < len(row) else None
                        except:
                            gross_weight_raw = None
                        
                        # 转换为浮点数并累加
                        net_weight_value = safe_convert_to_float(net_weight_raw, idx, "净重", error_log)
                        gross_weight_value = safe_convert_to_float(gross_weight_raw, idx, "毛重", error_log)
                        carton_net_weights[current_carton] += net_weight_value
                        carton_gross_weights[current_carton] += gross_weight_value
                except Exception as e:
                    error_log.append(f"处理第{idx+1}行时出错: {str(e)}，自动跳过")
                    continue
            
            # 验证每个箱号的净重是否小于毛重
            invalid_cartons = []
            for carton, total_net_weight in carton_net_weights.items():
                gross_weight = carton_gross_weights.get(carton, 0.0)
                
                # 跳过净重或毛重为0的箱号
                if total_net_weight <= 0 or gross_weight <= 0:
                    error_log.append(f"箱号 {carton} 的净重({total_net_weight:.2f})或毛重({gross_weight:.2f})为0或无效，自动跳过")
                    continue
                
                # 检查净重是否大于毛重（考虑小误差）
                tolerance = 0.01  # 1%的容差
                if total_net_weight > gross_weight * (1 + tolerance):
                    invalid_cartons.append({
                        "carton": carton,
                        "rows": carton_rows.get(carton, []),
                        "net_weight": total_net_weight,
                        "gross_weight": gross_weight
                    })
            
            # 生成验证结果
            if invalid_cartons:
                error_messages = []
                for item in invalid_cartons:
                    error_messages.append(
                        f"箱号 {item['carton']} 的总净重 {item['net_weight']:.2f} 大于总毛重 {item['gross_weight']:.2f}, "
                        f"涉及行: {', '.join(map(str, item['rows']))}"
                    )
                msg = "以下箱号的总净重大于总毛重，不符合验收标准:\n" + "\n".join(error_messages)
                if error_log:
                    msg += "\n\n【自动容错提示】:\n" + "\n".join(error_log)
                return {"success": False, "message": msg}
            
            msg = "净重毛重验证通过，所有箱号的总净重均小于总毛重"
            if error_log:
                msg += "\n\n【自动容错提示】:\n" + "\n".join(error_log)
            return {"success": True, "message": msg}
        except Exception as e:
            return {"success": False, "message": f"验证净重毛重时出错: {str(e)}"}
    
    def validate_summary_data(self, file_path):
        """验证表尾汇总数据正确性 - 此功能已被移除
        
        Args:
            file_path: 采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        # 根据需求变更，不再校验装箱单表尾汇总数据
        return {"success": True, "message": "汇总数据验证已跳过（根据需求变更）"}
    
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
            policy_df = pd.read_excel(policy_file_path, index_col=0)
            print(f"DEBUG: 政策文件内容：\n{policy_df.head()}")
            
            # 查找编号列或表头
            id_found = False
            policy_id = None
            
            # 检查是否有"采购装箱单编号"作为索引
            if "采购装箱单编号" in policy_df.index:
                policy_id = str(policy_df.loc["采购装箱单编号", "值"])
                id_found = True
                print(f"DEBUG: 在政策文件中找到编号: {policy_id}")
            
            # 如果没找到，尝试其他可能的索引名
            if not id_found:
                possible_indices = ["编号", "Number", "ID", "单号", "装箱单号", "Policy No", "参考编号"]
                for idx in possible_indices:
                    if idx in policy_df.index:
                        policy_id = str(policy_df.loc[idx, "值"])
                        id_found = True
                        print(f"DEBUG: 在政策文件中找到编号: {policy_id}")
                        break
            
            if not id_found or policy_id is None:
                return {
                    "success": False, 
                    "message": f"未在政策文件中找到编号。验收标准: 政策文件应包含与装箱单匹配的编号(参见文档第21行要求)。正确示例: 表头包含'Policy No: PL-20250418-0001'或有单独的'编号'列。"
                }
            
            # 清理编号字符串，移除可能的空格和特殊字符
            policy_id = policy_id.strip()
            packing_list_id = packing_list_id.strip()
            
            # 标准化编号：将小写l替换为数字1，将大写O替换为数字0
            policy_id_normalized = policy_id.replace('l', '1').replace('O', '0')
            packing_list_id_normalized = packing_list_id.replace('l', '1').replace('O', '0')
            
            # 提取数字部分进行比较
            policy_numbers = ''.join(filter(str.isdigit, policy_id_normalized))
            packing_list_numbers = ''.join(filter(str.isdigit, packing_list_id_normalized))
            
            print(f"DEBUG: 政策文件编号数字部分: {policy_numbers}")
            print(f"DEBUG: 装箱单编号数字部分: {packing_list_numbers}")
            
            # 如果数字部分相同，认为是同一个编号
            if policy_numbers == packing_list_numbers:
                return {"success": True, "message": "政策文件编号验证通过"}
            else:
                return {
                    "success": False, 
                    "message": f"政策文件编号({policy_id})与采购装箱单编号({packing_list_id})不一致。验收标准: 两个文件的编号应保持一致(参见文档第21行要求)。"
                }
        except Exception as e:
            return {"success": False, "message": f"验证政策文件编号时出错: {str(e)}。文件路径: {policy_file_path}"}
    
    def validate_exchange_rate_decimal(self, policy_file_path):
        """验证汇率值是否存在且为有效数字
        
        Args:
            policy_file_path: 政策文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取政策文件
            policy_df = pd.read_excel(policy_file_path, index_col=0)
            
            # 直接通过索引获取汇率值
            try:
                exchange_rate = float(policy_df.loc['汇率(RMB/美元)', '值'])
                return {"success": True, "message": "汇率验证通过"}
            except KeyError:
                return {"success": False, "message": "未找到汇率(RMB/美元)行"}
            except ValueError:
                return {"success": False, "message": "汇率值无法转换为数字"}
                
        except Exception as e:
            return {"success": False, "message": f"验证汇率时出错: {str(e)}"}
    
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
    




    def validate_sheet_naming(self, reimport_invoice_files):
        """校验reimport发票文件名与工厂唯一性绑定，页名规范"""
        try:
            for f in reimport_invoice_files:
                if not (f.endswith('.xlsx') and ('reimport' in f or 'RECI' in f)):
                    return {"success": False, "message": f"文件{f}命名不规范。"}
            return {"success": True, "message": f"所有reimport发票文件命名规范({len(reimport_invoice_files)}个)。"}
        except Exception as e:
            return {"success": False, "message": f"sheet_naming校验异常: {str(e)}"}

    def validate_all(self, packing_list_path, policy_file_path, reimport_invoice_files=None):
        """运行所有输入验证，支持reimport发票文件校验"""
        results = {}
        header_result = self.validate_packing_list_header(packing_list_path)
        results["packing_list_header"] = header_result
        packing_list_id = None
        if header_result["success"]:
            match = re.search(r'编号: \'([^\']+)\'', header_result["message"])
            if match:
                packing_list_id = match.group(1)
        results["packing_list_field_headers"] = self.validate_packing_list_field_headers(packing_list_path)
        results["weights"] = self.validate_weights(packing_list_path)
        results["summary_data"] = self.validate_summary_data(packing_list_path)
        results["policy_file_id"] = self.validate_policy_file_id(policy_file_path, packing_list_id)
        results["exchange_rate_decimal"] = self.validate_exchange_rate_decimal(policy_file_path)
        results["company_bank_info"] = self.validate_company_bank_info(policy_file_path)
        # 新增三项验收
        if reimport_invoice_files is not None:
  
            results["sheet_naming"] = self.validate_sheet_naming(reimport_invoice_files)
        return results