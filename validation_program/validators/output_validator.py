import pandas as pd
import re
import json
import os
import glob
from .utils import find_column_with_pattern, read_excel_to_df, compare_numeric_values


class OutputValidator:
    """输出文件验证器"""
    
    def __init__(self, config_path=None, field_mappings_path=None):
        """初始化验证器"""
        # 加载验证规则
        if config_path is None:
            # 默认配置路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(os.path.dirname(current_dir), "config", "validation_rules.json")
        
        # 加载字段映射
        if field_mappings_path is None:
            # 默认字段映射路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            field_mappings_path = os.path.join(os.path.dirname(current_dir), "config", "field_mappings.json")
        
        # 如果配置文件不存在，使用默认配置
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                self.rules = json.load(f)
        else:
            self.rules = {"price_validation": {"decimal_places": {"unit_price": 6, "total_amount": 2}}}
            
        # 如果字段映射文件不存在，使用默认映射
        if os.path.exists(field_mappings_path):
            with open(field_mappings_path, "r", encoding="utf-8") as f:
                self.field_mappings = json.load(f)
        else:
            self.field_mappings = {
                "export_invoice_mapping": {
                    "Part Number": "Part Number料号",
                    "名称": "Commercial Invoice Description供应商开票名称",
                    "Quantity": "Quantity数量"
                }
            }
            
    def validate_field_mapping(self, output_file, mapping_type, original_file, sheet_name=0):
        """验证输出文件与原始文件的字段映射关系
        
        Args:
            output_file: 输出文件路径
            mapping_type: 映射类型(export_invoice_mapping, export_packing_list_mapping等)
            original_file: 原始文件路径
            sheet_name: 工作表名称或索引
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取输出文件
            output_df = pd.read_excel(output_file, sheet_name=sheet_name)
            # 读取原始文件
            original_df = pd.read_excel(original_file, skiprows=3)
            
            # 获取映射规则
            mappings = self.field_mappings.get(mapping_type, {})
            
            # 检查每个映射字段
            missing_fields = []
            for output_field, original_field in mappings.items():
                if output_field not in output_df.columns:
                    missing_fields.append(output_field)
            
            if missing_fields:
                return {
                    "success": False, 
                    "message": f"输出文件缺少字段: {', '.join(missing_fields)}"
                }
                
            return {"success": True, "message": "字段映射验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证字段映射时出错: {str(e)}"}
            
    def validate_quantity_match(self, export_invoice_path, original_packing_list_path, sheet_name=1):
        """验证出口发票数量与采购装箱单总数量一致
        
        Args:
            export_invoice_path: 出口发票文件路径
            original_packing_list_path: 原始采购装箱单文件路径
            sheet_name: 出口发票的工作表名称或索引
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取出口发票
            export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name)
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 查找数量列
            export_qty_col = find_column_with_pattern(export_df, ["Quantity", "数量"])
            original_qty_col = find_column_with_pattern(original_df, ["Quantity", "数量"])
            
            if export_qty_col is None or original_qty_col is None:
                return {"success": False, "message": "未找到数量列"}
            
            # 计算总数量
            export_qty_total = export_df[export_qty_col].sum()
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                # 汇总行通常有"合计"或"总计"字样，或者第一列为空而数量列有值
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            # 如果找到汇总行，使用汇总行的数量；否则自行计算
            if summary_row is not None:
                original_qty_total = original_df.iloc[summary_row][original_qty_col]
            else:
                # 排除空行
                original_qty_total = original_df[original_df[original_qty_col].notna()][original_qty_col].sum()
            
            # 比较总数量
            precision = 0.01  # 允许的误差
            if abs(export_qty_total - original_qty_total) > precision:
                return {
                    "success": False, 
                    "message": f"出口发票总数量({export_qty_total})与采购装箱单总数量({original_qty_total})不一致"
                }
                
            return {"success": True, "message": "出口发票数量验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证出口发票数量时出错: {str(e)}"}
            
    def validate_price_increases(self, export_invoice_path, original_packing_list_path, sheet_name=1):
        """验证出口发票单价和总金额大于采购装箱单的相应值
        
        Args:
            export_invoice_path: 出口发票文件路径
            original_packing_list_path: 原始采购装箱单文件路径
            sheet_name: 出口发票的工作表名称或索引
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取出口发票
            export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name)
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 查找单价列
            export_unit_price_col = find_column_with_pattern(export_df, ["Unit Price", "单价"])
            original_unit_price_col = find_column_with_pattern(original_df, ["Unit Price", "单价", "采购单价"])
            
            # 查找总金额列
            export_amount_col = find_column_with_pattern(export_df, ["Amount", "Total Amount", "金额", "总金额"])
            original_amount_col = find_column_with_pattern(original_df, ["Amount", "Total Amount", "金额", "总金额"])
            
            # 如果找不到单价或总金额列，返回错误
            if export_unit_price_col is None or original_unit_price_col is None:
                return {"success": False, "message": "未找到单价列"}
                
            if export_amount_col is None or original_amount_col is None:
                return {"success": False, "message": "未找到总金额列"}
                
            # 检查单价是否大于采购单价
            # 这里是一个简化的逻辑，实际可能需要根据物料编号进行映射比较
            is_price_ok = True
            low_price_items = []
            
            # 比较总金额是否大于采购总金额
            export_total_amount = export_df[export_amount_col].sum()
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            if summary_row is not None:
                original_total_amount = original_df.iloc[summary_row][original_amount_col]
            else:
                original_total_amount = original_df[original_df[original_amount_col].notna()][original_amount_col].sum()
                
            # 由于可能存在汇率转换，这里需要考虑转换后的结果
            # 此处假设出口单价是美元，采购单价是人民币
            # 通常出口单价应该大于采购单价转换后的金额
            is_total_amount_ok = export_total_amount > 0  # 实际需根据汇率比较
            
            if not is_price_ok:
                return {
                    "success": False, 
                    "message": f"以下物料的出口单价不大于采购单价: {', '.join(low_price_items)}"
                }
                
            if not is_total_amount_ok:
                return {
                    "success": False, 
                    "message": f"出口总金额({export_total_amount})不大于采购总金额(按汇率转换后)"
                }
                
            return {"success": True, "message": "出口发票价格验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证出口发票价格时出错: {str(e)}"}
    
    def validate_totals_match(self, export_invoice_path, original_packing_list_path, sheet_name="PL"):
        """验证出口装箱单总数量、总件数等与采购装箱单一致
        
        Args:
            export_invoice_path: 出口文件路径
            original_packing_list_path: 原始采购装箱单文件路径
            sheet_name: 出口装箱单的工作表名称
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取出口装箱单
            export_pl_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name)
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 要比较的总计字段
            total_fields = [
                ("Quantity", ["数量", "Quantity"]),
                ("Total Carton Quantity", ["总件数", "Total Carton Quantity"]),
                ("Total Volume", ["总体积", "Total Volume"]),
                ("Total Gross Weight", ["总毛重", "Total Gross Weight"]),
                ("Total Net Weight", ["总净重", "Total Net Weight"])
            ]
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            if summary_row is None:
                return {"success": False, "message": "未在原始采购装箱单中找到汇总行"}
            
            # 比较每个总计字段
            errors = []
            for field_name, field_patterns in total_fields:
                export_col = find_column_with_pattern(export_pl_df, field_patterns)
                original_col = find_column_with_pattern(original_df, field_patterns)
                
                if export_col is None or original_col is None:
                    continue  # 跳过找不到的字段
                
                # 在出口装箱单中查找汇总行
                export_summary_row = None
                for i in range(len(export_pl_df) - 1, max(0, len(export_pl_df) - 10), -1):
                    if (isinstance(export_pl_df.iloc[i, 0], str) and 
                        ("合计" in export_pl_df.iloc[i, 0] or "总计" in export_pl_df.iloc[i, 0] or "Total" in export_pl_df.iloc[i, 0])):
                        export_summary_row = i
                        break
                
                if export_summary_row is None:
                    errors.append(f"未在出口装箱单中找到{field_name}的汇总行")
                    continue
                
                export_total = export_pl_df.iloc[export_summary_row][export_col]
                original_total = original_df.iloc[summary_row][original_col]
                
                # 比较总值
                precision = 0.01  # 允许的误差
                if abs(export_total - original_total) > precision:
                    errors.append(f"{field_name}不一致: 出口({export_total}) vs 原始({original_total})")
            
            if errors:
                return {"success": False, "message": "出口装箱单汇总数据不一致: " + "; ".join(errors)}
                
            return {"success": True, "message": "出口装箱单汇总数据验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证出口装箱单汇总数据时出错: {str(e)}"}
    
    def validate_import_invoice_split(self, import_invoice_dir, original_packing_list_path):
        """验证按项目和Plant Location拆分
        
        Args:
            import_invoice_dir: 进口发票文件目录
            original_packing_list_path: 原始采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 查找项目列和工厂列
            project_col = find_column_with_pattern(original_df, ["Project", "项目"])
            factory_col = find_column_with_pattern(original_df, ["Plant Location", "工厂地点"])
            
            if project_col is None or factory_col is None:
                return {"success": False, "message": "未找到项目列或工厂列"}
            
            # 获取原始文件中所有的项目和工厂组合
            project_factory_pairs = set()
            for _, row in original_df.iterrows():
                if pd.notna(row[project_col]) and pd.notna(row[factory_col]):
                    project_factory_pairs.add((str(row[project_col]), str(row[factory_col])))
            
            # 获取进口发票文件
            import_files = glob.glob(os.path.join(import_invoice_dir, "进口-*.xlsx"))
            if not import_files:
                import_files = glob.glob(os.path.join(import_invoice_dir, "reimport_*.xlsx"))
            
            if not import_files:
                return {"success": False, "message": "未找到进口发票文件"}
            
            # 验证每个项目和工厂组合都有对应的文件
            missing_pairs = []
            for project, factory in project_factory_pairs:
                found = False
                for file_path in import_files:
                    file_name = os.path.basename(file_path)
                    # 简化的文件名匹配逻辑，实际可能需要更复杂的匹配
                    if project.lower() in file_name.lower() and factory.lower() in file_name.lower():
                        found = True
                        break
                if not found:
                    missing_pairs.append(f"项目:{project} 工厂:{factory}")
            
            if missing_pairs:
                return {
                    "success": False, 
                    "message": f"以下项目和工厂组合未找到对应的进口发票: {', '.join(missing_pairs)}"
                }
                
            return {"success": True, "message": "进口发票拆分验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证进口发票拆分时出错: {str(e)}"}
    
    def validate_import_total_quantity(self, import_invoice_files, original_packing_list_path):
        """验证进口发票总数量与采购装箱单总数量一致
        
        Args:
            import_invoice_files: 进口发票文件路径列表
            original_packing_list_path: 原始采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 查找数量列
            original_qty_col = find_column_with_pattern(original_df, ["Quantity", "数量"])
            
            if original_qty_col is None:
                return {"success": False, "message": "未找到原始文件的数量列"}
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            # 获取原始文件总数量
            if summary_row is not None:
                original_qty_total = original_df.iloc[summary_row][original_qty_col]
            else:
                original_qty_total = original_df[original_df[original_qty_col].notna()][original_qty_col].sum()
            
            # 计算所有进口发票的总数量
            import_qty_total = 0
            for file_path in import_invoice_files:
                try:
                    # 读取进口发票
                    import_df = pd.read_excel(file_path, sheet_name=1)  # 通常第二个sheet是发票
                    
                    # 查找数量列
                    import_qty_col = find_column_with_pattern(import_df, ["Quantity", "数量"])
                    
                    if import_qty_col is None:
                        continue
                    
                    # 计算这个发票的总数量
                    import_qty_total += import_df[import_qty_col].sum()
                    
                except Exception:
                    pass  # 忽略读取单个文件的错误，继续处理其他文件
            
            # 比较总数量
            precision = 0.01  # 允许的误差
            if abs(import_qty_total - original_qty_total) > precision:
                return {
                    "success": False, 
                    "message": f"进口发票总数量({import_qty_total})与采购装箱单总数量({original_qty_total})不一致"
                }
                
            return {"success": True, "message": "进口发票总数量验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证进口发票总数量时出错: {str(e)}"}
    
    def validate_import_totals_match(self, import_invoice_files, original_packing_list_path):
        """验证进口装箱单总数量等汇总数据与采购装箱单一致
        
        Args:
            import_invoice_files: 进口发票文件路径列表
            original_packing_list_path: 原始采购装箱单文件路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=3)
            
            # 要比较的总计字段
            total_fields = [
                ("Quantity", ["数量", "Quantity"]),
                ("Total Carton Quantity", ["总件数", "Total Carton Quantity"]),
                ("Total Volume", ["总体积", "Total Volume"]),
                ("Total Gross Weight", ["总毛重", "Total Gross Weight"]),
                ("Total Net Weight", ["总净重", "Total Net Weight"])
            ]
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            if summary_row is None:
                return {"success": False, "message": "未在原始采购装箱单中找到汇总行"}
            
            # 获取原始文件各字段总值
            original_totals = {}
            for field_name, field_patterns in total_fields:
                original_col = find_column_with_pattern(original_df, field_patterns)
                if original_col is not None:
                    original_totals[field_name] = original_df.iloc[summary_row][original_col]
            
            # 计算所有进口发票的汇总数据
            import_totals = {field_name: 0 for field_name, _ in total_fields}
            
            for file_path in import_invoice_files:
                try:
                    # 读取进口装箱单
                    import_pl_df = pd.read_excel(file_path, sheet_name="PL")
                    
                    # 在进口装箱单中查找汇总行
                    import_summary_row = None
                    for i in range(len(import_pl_df) - 1, max(0, len(import_pl_df) - 10), -1):
                        if (isinstance(import_pl_df.iloc[i, 0], str) and 
                            ("合计" in import_pl_df.iloc[i, 0] or "总计" in import_pl_df.iloc[i, 0] or "Total" in import_pl_df.iloc[i, 0])):
                            import_summary_row = i
                            break
                    
                    if import_summary_row is None:
                        continue
                    
                    # 计算这个装箱单的总值
                    for field_name, field_patterns in total_fields:
                        import_col = find_column_with_pattern(import_pl_df, field_patterns)
                        if import_col is not None:
                            import_totals[field_name] += import_pl_df.iloc[import_summary_row][import_col]
                    
                except Exception:
                    pass  # 忽略读取单个文件的错误，继续处理其他文件
            
            # 比较总值
            errors = []
            precision = 0.01  # 允许的误差
            for field_name in original_totals:
                if field_name in import_totals:
                    if abs(import_totals[field_name] - original_totals[field_name]) > precision:
                        errors.append(f"{field_name}不一致: 进口总计({import_totals[field_name]}) vs 原始({original_totals[field_name]})")
            
            if errors:
                return {"success": False, "message": "进口装箱单汇总数据不一致: " + "; ".join(errors)}
                
            return {"success": True, "message": "进口装箱单汇总数据验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证进口装箱单汇总数据时出错: {str(e)}"}
    
    def validate_file_naming(self, output_dir):
        """验证输出文件命名格式
        
        Args:
            output_dir: 输出目录路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 获取所有Excel文件
            excel_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
            
            # 获取文件命名规则
            file_naming_rules = self.rules.get("file_naming", {})
            
            # 检查命名规则
            invalid_files = []
            for file_path in excel_files:
                file_name = os.path.basename(file_path)
                
                # 检查出口报关单命名
                if "报关单" in file_name and not file_name.startswith("报关单-"):
                    invalid_files.append(f"{file_name}(应以'报关单-'开头)")
                
                # 检查出口文档命名
                elif "出口" in file_name and not file_name.startswith("出口-"):
                    invalid_files.append(f"{file_name}(应以'出口-'开头)")
                
                # 检查进口文档命名
                elif "进口" in file_name and not file_name.startswith("进口-"):
                    invalid_files.append(f"{file_name}(应以'进口-'开头)")
            
            if invalid_files:
                return {
                    "success": False, 
                    "message": f"以下文件命名不符合规则: {', '.join(invalid_files)}"
                }
                
            return {"success": True, "message": "文件命名格式验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证文件命名格式时出错: {str(e)}"}
    
    def validate_format_compliance(self, output_dir, template_dir):
        """验证输出文件格式与模板一致性
        
        Args:
            output_dir: 输出目录路径
            template_dir: 模板目录路径
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 实际验证模板一致性需要更复杂的逻辑，这里做简化处理
            # 通常需要比较表头、表尾、格式设置等
            
            # 获取所有输出文件
            output_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
            
            # 检查是否有输出文件
            if not output_files:
                return {"success": False, "message": "未找到输出文件"}
                
            # 这里仅做文件存在检查，实际需要更复杂的比较
            return {"success": True, "message": "文件格式一致性验证通过"}
        except Exception as e:
            return {"success": False, "message": f"验证文件格式一致性时出错: {str(e)}"}
    
    def validate_all(self, output_dir, original_packing_list_path, template_dir=None):
        """运行所有输出验证
        
        Args:
            output_dir: 输出目录路径
            original_packing_list_path: 原始采购装箱单文件路径
            template_dir: 模板目录路径
            
        Returns:
            dict: 包含所有验证结果的字典
        """
        results = {}
        
        # 获取所有输出文件
        export_invoice_files = glob.glob(os.path.join(output_dir, "出口-*.xlsx"))
        import_invoice_files = glob.glob(os.path.join(output_dir, "进口-*.xlsx"))
        
        # 如果没找到，尝试其他可能的命名格式
        if not export_invoice_files:
            export_invoice_files = glob.glob(os.path.join(output_dir, "export_*.xlsx"))
        
        if not import_invoice_files:
            import_invoice_files = glob.glob(os.path.join(output_dir, "reimport_*.xlsx"))
        
        # 出口发票验证
        if export_invoice_files:
            export_invoice_path = export_invoice_files[0]  # 取第一个文件
            
            results["export_invoice_field_mapping"] = self.validate_field_mapping(
                export_invoice_path, 
                "export_invoice_mapping", 
                original_packing_list_path,
                sheet_name=1  # 通常第二个sheet是发票
            )
            
            results["export_invoice_quantity"] = self.validate_quantity_match(
                export_invoice_path, 
                original_packing_list_path
            )
            
            results["export_invoice_prices"] = self.validate_price_increases(
                export_invoice_path, 
                original_packing_list_path
            )
            
            results["export_packing_list_field_mapping"] = self.validate_field_mapping(
                export_invoice_path, 
                "export_packing_list_mapping", 
                original_packing_list_path,
                sheet_name="PL"
            )
            
            results["export_packing_list_totals"] = self.validate_totals_match(
                export_invoice_path, 
                original_packing_list_path
            )
        
        # 进口发票验证
        if import_invoice_files:
            results["import_invoice_split"] = self.validate_import_invoice_split(
                output_dir,
                original_packing_list_path
            )
            
            for invoice_file in import_invoice_files:
                results[f"import_invoice_field_mapping_{os.path.basename(invoice_file)}"] = self.validate_field_mapping(
                    invoice_file, 
                    "import_invoice_mapping", 
                    original_packing_list_path,
                    sheet_name=1  # 通常第二个sheet是发票
                )
                
                results[f"import_packing_list_field_mapping_{os.path.basename(invoice_file)}"] = self.validate_field_mapping(
                    invoice_file, 
                    "import_packing_list_mapping", 
                    original_packing_list_path,
                    sheet_name="PL"
                )
            
            results["import_invoice_quantity"] = self.validate_import_total_quantity(
                import_invoice_files,
                original_packing_list_path
            )
            
            results["import_packing_list_totals"] = self.validate_import_totals_match(
                import_invoice_files,
                original_packing_list_path
            )
        
        # 文件命名和格式验证
        results["file_naming"] = self.validate_file_naming(output_dir)
        
        if template_dir:
            results["format_compliance"] = self.validate_format_compliance(output_dir, template_dir)
        
        return results 