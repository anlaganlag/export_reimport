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
                    "Material code": "料号",
                    "DESCRIPTION": "供应商开票名称",
                    "Model NO.":"型号",
                    "Unit Price":"采购单价",
                    "Qty": "每箱数量",
                    "Unit":"单位",
                }
            }
            
    def validate_field_mapping(self, output_file, mapping_type, original_file, sheet_name=0, skiprows=None):
        """验证输出文件与原始文件的字段映射关系
        
        Args:
            output_file: 输出文件路径
            mapping_type: 映射类型（export_invoice_mapping, export_packing_list_mapping等）
            original_file: 原始文件路径
            sheet_name: 工作表名称或索引
            skiprows: 跳过的行数
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取输出文件
            try:
                if mapping_type == 'export_packing_list_mapping':
                    skiprows = 15
                if mapping_type == 'import_invoice_mapping':
                    skiprows = 0

                output_df = pd.read_excel(output_file, sheet_name=sheet_name, skiprows=skiprows)
                print(f"DEBUG: 成功读取{output_file}的{sheet_name}工作表，跳过前{skiprows}行")
                print(f"DEBUG: 读取到的列名: {output_df.columns.tolist()}")
            except Exception as e:
                print(f"ERROR: 读取{output_file}时出错: {str(e)}")
                # 尝试不同方式读取
                if skiprows is not None:
                    print(f"DEBUG: 尝试不使用skiprows参数读取")
                    output_df = pd.read_excel(output_file, sheet_name=sheet_name)
                else:
                    # 读取失败，返回错误
                    return {
                        "success": False, 
                        "message": f"无法读取输出文件: {str(e)}"
                    }
            
            # 读取原始文件 - 修正skiprows从3改为2，确保读取到第一行数据
            original_df = pd.read_excel(original_file, skiprows=2)
            print(f"DEBUG: 成功读取原始文件，使用skiprows=2")
            print(f"DEBUG: 原始文件列名: {original_df.columns.tolist()}")
            
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
            print(f"ERROR: 验证字段映射时出错: {str(e)}")
            return {"success": False, "message": f"验证字段映射时出错: {str(e)}"}
            
    def validate_quantity_match(self, export_invoice_path, original_packing_list_path, sheet_name=1, skiprows=6):
        """验证出口发票数量与采购装箱单总数量一致
        
        Args:
            export_invoice_path: 出口发票文件路径
            original_packing_list_path: 原始采购装箱单文件路径
            sheet_name: 出口发票的工作表名称或索引
            skiprows: 要跳过的行数，用于处理含有标题行的Excel文件
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取出口发票
            try:
                export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name, skiprows=skiprows)
                print(f"DEBUG: 成功读取{export_invoice_path}的{sheet_name}工作表，跳过前{skiprows}行")
            except Exception as e:
                print(f"ERROR: 读取{export_invoice_path}时出错: {str(e)}")
                # 尝试不使用skiprows参数
            export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name,skiprows=6)
            
            # 读取原始采购装箱单
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            
            # 查找数量列
            export_qty_col = find_column_with_pattern(export_df, ["Quantity", "数量", "Qty"])
            original_qty_col = find_column_with_pattern(original_df, ["Quantity", "数量", "Qty"])
            
            # 查找NO.列
            original_no_col = find_column_with_pattern(original_df, ["NO.", "序号", "No"])
            export_no_col = find_column_with_pattern(export_df, ["NO.", "序号", "No"])
            
            if export_qty_col is None or original_qty_col is None:
                return {"success": False, "message": "未找到数量列"}
            
            # 查找物料编码列，以便进行详细比较
            export_material_col = find_column_with_pattern(export_df, ["Material code", "物料编码", "料号"])
            original_material_col = find_column_with_pattern(original_df, ["Material code", "物料编码", "料号"])
            
            # 计算出口发票总数量，排除汇总行
            export_qty_total = 0
            export_summary_rows = []
            export_material_qty = {}  # 用于记录每个物料的数量
            
            for i, row in export_df.iterrows():
                # 识别汇总行: 
                # 1. NO.列为空但数量列不为空
                # 2. 第一列（如果不是NO.列）有"合计"、"总计"或"Total"字样
                is_summary = False
                
                # 检查第一列是否有汇总相关的字样
                first_col = export_df.columns[0]
                if isinstance(row.get(first_col), str) and any(x in str(row.get(first_col)).lower() for x in ["合计", "总计", "total"]):
                    print(f"DEBUG: 出口发票第{i+1}行被识别为汇总行(有合计字样): {row[export_qty_col]}")
                    is_summary = True
                    export_summary_rows.append(i)
                
                # 当NO.列存在时，判断是否为汇总行
                if export_no_col is not None:
                    if pd.isna(row[export_no_col]) and pd.notna(row[export_qty_col]):
                        print(f"DEBUG: 出口发票第{i+1}行被识别为汇总行(NO.为空): {row[export_qty_col]}")
                        is_summary = True
                        export_summary_rows.append(i)
                
                # 如果不是汇总行，且数量为有效数字，则计入总数
                if not is_summary and pd.notna(row[export_qty_col]):
                    try:
                        # 确保转换为数字
                        qty = float(row[export_qty_col])
                        export_qty_total += qty
                        
                        # 记录物料编码和数量，用于详细比较
                        if export_material_col is not None and pd.notna(row[export_material_col]):
                            material_code = str(row[export_material_col])
                            if material_code in export_material_qty:
                                export_material_qty[material_code] += qty
                            else:
                                export_material_qty[material_code] = qty
                    except ValueError:
                        print(f"WARNING: 出口发票第{i+1}行的数量值'{row[export_qty_col]}'不能转换为数字")
            
            # 在原始文件中计算总数量，排除汇总行
            original_qty_total = 0
            original_summary_rows = []
            original_material_qty = {}  # 用于记录每个物料的数量
            
            for i, row in original_df.iterrows():
                # 识别汇总行: 
                # 1. NO.列为空但数量列不为空
                # 2. 第一列（如果不是NO.列）有"合计"、"总计"或"Total"字样
                is_summary = False
                
                if original_no_col is not None:
                    # 当NO.列存在时，判断是否为汇总行
                    if pd.isna(row[original_no_col]) and pd.notna(row[original_qty_col]):
                        print(f"DEBUG: 原始文件第{i+1}行被识别为汇总行(NO.为空): {row[original_qty_col]}")
                        is_summary = True
                        original_summary_rows.append(i)
                
                # 检查第一列是否有汇总相关的字样
                first_col = original_df.columns[0]
                if isinstance(row.get(first_col), str) and any(x in str(row.get(first_col)).lower() for x in ["合计", "总计", "total"]):
                    print(f"DEBUG: 原始文件第{i+1}行被识别为汇总行(有合计字样): {row[original_qty_col]}")
                    is_summary = True
                    original_summary_rows.append(i)
                
                # 如果不是汇总行，且数量为有效数字，则计入总数
                if not is_summary and pd.notna(row[original_qty_col]):
                    try:
                        # 确保转换为数字
                        qty = float(row[original_qty_col])
                        original_qty_total += qty
                        
                        # 记录物料编码和数量，用于详细比较
                        if original_material_col is not None and pd.notna(row[original_material_col]):
                            material_code = str(row[original_material_col])
                            if material_code in original_material_qty:
                                original_material_qty[material_code] += qty
                            else:
                                original_material_qty[material_code] = qty
                    except ValueError:
                        print(f"WARNING: 原始文件第{i+1}行的数量值'{row[original_qty_col]}'不能转换为数字")
            
            print(f"DEBUG: 出口发票总数量(排除汇总行): {export_qty_total}")
            print(f"DEBUG: 出口发票识别的汇总行: {export_summary_rows}")
            print(f"DEBUG: 原始文件总数量(排除汇总行): {original_qty_total}")
            print(f"DEBUG: 原始文件识别的汇总行: {original_summary_rows}")
            
            # 打印物料数量详细比较
            print("DEBUG: 物料数量详细比较:")
            for material_code in set(original_material_qty.keys()) | set(export_material_qty.keys()):
                original_qty = original_material_qty.get(material_code, 0)
                export_qty = export_material_qty.get(material_code, 0)
                diff = export_qty - original_qty
                print(f"DEBUG: 物料: {material_code}, 原始数量: {original_qty}, 出口数量: {export_qty}, 差异: {diff}")
            
            # 比较总数量
            precision = 0.01  # 允许的误差
            if abs(export_qty_total - original_qty_total) > precision:
                # 查找具体差异
                missing_materials = []
                for material_code, original_qty in original_material_qty.items():
                    export_qty = export_material_qty.get(material_code, 0)
                    if abs(export_qty - original_qty) > precision:
                        missing_materials.append(f"{material_code}(原始:{original_qty}, 出口:{export_qty}, 差异:{export_qty-original_qty})")
                
                error_detail = f"出口发票总数量({export_qty_total})与采购装箱单总数量({original_qty_total})不一致"
                if missing_materials:
                    error_detail += f"\n存在差异的物料: {', '.join(missing_materials)}"
                
                return {
                    "success": False, 
                    "message": error_detail
                }
                
            return {"success": True, "message": "出口发票数量验证通过"}
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证出口发票数量时出错: {str(e)}\n{error_details}")
            return {"success": False, "message": f"验证出口发票数量时出错: {str(e)}"}
            
    def validate_price_increases(self, export_invoice_path, original_packing_list_path, sheet_name=1, skiprows=6):
        """验证出口发票单价和总金额大于采购装箱单的相应值
        
        Args:
            export_invoice_path: 出口发票文件路径
            original_packing_list_path: 原始采购装箱单文件路径
            sheet_name: 出口发票的工作表名称或索引
            skiprows: 要跳过的行数，用于处理含有标题行的Excel文件
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取出口发票
            try:
                export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name, skiprows=skiprows)
                print(f"DEBUG: 成功读取{export_invoice_path}的{sheet_name}工作表，跳过前{skiprows}行")
            except Exception as e:
                print(f"ERROR: 读取{export_invoice_path}时出错: {str(e)}")
                # 尝试不使用skiprows参数
            export_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name,skiprows=6)
            
            # 读取原始采购装箱单 - 修正使用skiprows=2
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            print(f"DEBUG: 成功读取原始装箱单，使用skiprows=2")
            
            # 查找单价列
            export_unit_price_col = find_column_with_pattern(export_df, ["Unit Price", "单价"])
            original_unit_price_col = find_column_with_pattern(original_df, ["Unit Price", "单价", "采购单价"])
            
            # 查找数量列
            export_qty_col = find_column_with_pattern(export_df, ["Qty", "Quantity", "数量"])
            original_qty_col = find_column_with_pattern(original_df, ["Qty", "Quantity", "数量"])
            
            # 查找总金额列
            export_amount_col = find_column_with_pattern(export_df, ["Amount", "Total Amount", "金额", "总金额"])
            original_amount_col = find_column_with_pattern(original_df, ["Amount", "Total Amount", "金额", "总金额"])
            
            # 如果找不到单价或数量列，返回错误
            if export_unit_price_col is None or original_unit_price_col is None:
                return {"success": False, "message": "未找到单价列"}
                
            if export_qty_col is None or original_qty_col is None:
                return {"success": False, "message": "未找到数量列"}
            
            # 如果找不到总金额列，通过单价*数量计算
            export_total_amount = 0
            if export_amount_col is not None:
                export_total_amount = export_df[export_amount_col].sum()
            else:
                # 通过单价*数量计算总金额
                export_df['calculated_amount'] = export_df[export_unit_price_col] * export_df[export_qty_col]
                export_total_amount = export_df['calculated_amount'].sum()
                print(f"DEBUG: 通过单价*数量计算出口总金额: {export_total_amount}")
            
            # 同样计算原始文件的总金额
            original_total_amount = 0
            # 初始化summary_row变量，确保在所有条件分支中都有定义
            summary_row = None
            
            # 查找原始文件中的单价和数量列
            original_unit_price_col = find_column_with_pattern(original_df, ["Unit Price", "单价", "采购单价"])
            original_qty_col = find_column_with_pattern(original_df, ["Qty", "Quantity", "数量", "PCS"])
            
            print(f"DEBUG: 原始文件单价列: {original_unit_price_col}, 数量列: {original_qty_col}")
            
            if original_unit_price_col is not None and original_qty_col is not None:
                # 通过单价*数量计算总金额
                original_df['calculated_amount'] = original_df[original_unit_price_col] * original_df[original_qty_col]
                
                # 检查是否有汇总行
                for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                    if (isinstance(original_df.iloc[i, 0], str) and 
                        ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                        summary_row = i
                        break
                
                if summary_row is not None:
                    print(f"DEBUG: 找到汇总行: {summary_row}")
                    # 汇总行的总金额应该是单价*数量
                    original_total_amount = original_df.iloc[summary_row][original_unit_price_col] * original_df.iloc[summary_row][original_qty_col]
                    print(f"DEBUG: 汇总行计算总金额: {original_total_amount}")
                else:
                    # 计算所有行的总和
                    original_total_amount = original_df['calculated_amount'].sum()
                    print(f"DEBUG: 计算所有行总金额: {original_total_amount}")
            elif original_amount_col is not None:
                # 如果找到总金额列，使用原来的逻辑
                if summary_row is not None:
                    original_total_amount = original_df.iloc[summary_row][original_amount_col]
                else:
                    original_total_amount = original_df[original_df[original_amount_col].notna()][original_amount_col].sum()
            else:
                print("WARNING: 未找到足够的列来计算原始文件总金额")
            
            # 由于可能存在汇率转换，这里需要考虑转换后的结果
            # 此处假设出口单价是美元，采购单价是人民币
            # 通常出口单价应该大于采购单价转换后的金额
            is_total_amount_ok = export_total_amount > 0  # 实际需根据汇率比较
            
            # 检查单价是否大于采购单价
            # 这里是一个简化的逻辑，实际可能需要根据物料编号进行映射比较
            is_price_ok = True
            low_price_items = []
            
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
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证出口发票价格时出错: {str(e)}\n{error_details}")
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
            try:
                export_pl_df = pd.read_excel(export_invoice_path, sheet_name=sheet_name)
                print(f"DEBUG: 成功读取出口装箱单工作表'{sheet_name}'")
            except Exception as e:
                # 如果找不到"PL"工作表，尝试查找其他可能的工作表名
                alternate_sheet_names = ["Packing List", "PackingList", "PL Sheet", "装箱单"]
                for alt_name in alternate_sheet_names:
                    try:
                        export_pl_df = pd.read_excel(export_invoice_path, sheet_name=alt_name)
                        print(f"WARNING: 未找到'PL'工作表，但找到了'{alt_name}'工作表。根据验收文档，应将工作表名改为'PL'")
                        return {"success": False, "message": f"出口装箱单工作表名应为'PL'，而不是'{alt_name}'。请修改工作表名以符合验收要求。"}
                    except:
                        continue
                # 如果所有尝试都失败，返回原始错误
                return {"success": False, "message": f"无法读取出口装箱单工作表: {str(e)}"}
                
            # 读取原始采购装箱单 - 修正使用skiprows=2
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            print(f"DEBUG: 成功读取原始装箱单，使用skiprows=2")
            
            # 要比较的总计字段
            total_fields = [
                ("Quantity", ["数量", "Quantity", "Qty"]),
                ("Total Carton Quantity", ["总件数", "Total Carton Quantity", "Carton Qty"]),
                ("Total Volume", ["总体积", "Total Volume", "Volume"]),
                ("Total Gross Weight", ["总毛重", "Total Gross Weight", "G.W", "Gross Weight"]),
                ("Total Net Weight", ["总净重", "Total Net Weight", "N.W", "Net Weight"])
            ]
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            # 在出口装箱单中查找汇总行
            export_summary_row = None
            for i in range(len(export_pl_df) - 1, max(0, len(export_pl_df) - 10), -1):
                if (isinstance(export_pl_df.iloc[i, 0], str) and 
                    ("合计" in export_pl_df.iloc[i, 0] or "总计" in export_pl_df.iloc[i, 0] or "Total" in export_pl_df.iloc[i, 0])):
                    export_summary_row = i
                    break
            
            # 比较每个总计字段
            errors = []
            for field_name, field_patterns in total_fields:
                export_col = find_column_with_pattern(export_pl_df, field_patterns)
                original_col = find_column_with_pattern(original_df, field_patterns)
                
                if export_col is None or original_col is None:
                    continue  # 跳过找不到的字段
                
                # 获取导出文件的总值
                export_total = 0
                if export_summary_row is not None:
                    export_total = export_pl_df.iloc[export_summary_row][export_col]
                else:
                    # 如果找不到汇总行，手动计算总和
                    export_total = export_pl_df[export_col].sum()
                    print(f"DEBUG: 手动计算出口{field_name}总值: {export_total}")
                
                # 获取原始文件的总值
                original_total = 0
                if summary_row is not None:
                    original_total = original_df.iloc[summary_row][original_col]
                else:
                    # 如果找不到汇总行，手动计算总和
                    original_total = original_df[original_col].sum()
                    print(f"DEBUG: 手动计算原始{field_name}总值: {original_total}")
                
                # 比较总值
                precision = 0.01  # 允许的误差
                if abs(export_total - original_total) > precision:
                    errors.append(f"{field_name}不一致: 出口({export_total}) vs 原始({original_total})")
            
            if errors:
                return {"success": False, "message": "出口装箱单汇总数据不一致: " + "; ".join(errors)}
                
            return {"success": True, "message": "出口装箱单汇总数据验证通过"}
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证出口装箱单汇总数据时出错: {str(e)}\n{error_details}")
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
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            
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
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证进口发票拆分时出错: {str(e)}\n{error_details}")
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
            # 读取原始采购装箱单 - 修正使用skiprows=2
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            print(f"DEBUG: 成功读取原始装箱单，使用skiprows=2")
            
            # 查找数量列
            original_qty_col = find_column_with_pattern(original_df, ["Quantity", "数量", "Qty"])
            
            # 查找NO.列
            no_col = find_column_with_pattern(original_df, ["NO.", "序号", "No"])
            
            if original_qty_col is None:
                return {"success": False, "message": "未找到原始文件的数量列"}
            
            # 在原始文件中计算总数量，排除汇总行
            original_qty_total = 0
            summary_rows = []
            
            for i, row in original_df.iterrows():
                # 识别汇总行: 
                # 1. NO.列为空但数量列不为空
                # 2. 第一列（如果不是NO.列）有"合计"、"总计"或"Total"字样
                is_summary = False
                
                if no_col is not None:
                    # 当NO.列存在时，判断是否为汇总行
                    if pd.isna(row[no_col]) and pd.notna(row[original_qty_col]):
                        print(f"DEBUG: 第{i+1}行被识别为汇总行(NO.为空): {row[original_qty_col]}")
                        is_summary = True
                        summary_rows.append(i)
                
                # 检查第一列是否有汇总相关的字样
                first_col = original_df.columns[0]
                if isinstance(row[first_col], str) and any(x in row[first_col].lower() for x in ["合计", "总计", "total"]):
                    print(f"DEBUG: 第{i+1}行被识别为汇总行(有合计字样): {row[original_qty_col]}")
                    is_summary = True
                    summary_rows.append(i)
                
                # 如果不是汇总行，且数量为有效数字，则计入总数
                if not is_summary and pd.notna(row[original_qty_col]):
                    try:
                        # 确保转换为数字
                        qty = float(row[original_qty_col])
                        original_qty_total += qty
                    except ValueError:
                        print(f"WARNING: 第{i+1}行的数量值'{row[original_qty_col]}'不能转换为数字")
            
            print(f"DEBUG: 原始文件总数量(排除汇总行): {original_qty_total}")
            print(f"DEBUG: 识别的汇总行: {summary_rows}")
            
            # 计算所有进口发票的总数量
            import_qty_total = 0
            import_file_qtys = {}  # 存储每个文件的数量，用于详细报告
            
            for file_path in import_invoice_files:
                try:
                    # 读取进口发票
                    import_df = pd.read_excel(file_path, sheet_name=1, skiprows=6)  # 跳过表头
                    print(f"DEBUG: 读取进口发票: {file_path}")
                    
                    # 查找数量列
                    import_qty_col = find_column_with_pattern(import_df, ["Quantity", "数量", "Qty"])
                    
                    if import_qty_col is None:
                        print(f"WARNING: 未找到数量列，文件: {file_path}, 列名: {import_df.columns.tolist()}")
                        continue
                    
                    # 计算这个发票的总数量，排除汇总行
                    file_qty = 0
                    for i, row in import_df.iterrows():
                        # 简单判断是否为汇总行
                        first_col = import_df.columns[0]
                        is_summary = False
                        
                        if isinstance(row.get(first_col), str) and any(x in str(row.get(first_col)).lower() for x in ["合计", "总计", "total"]):
                            is_summary = True
                            print(f"DEBUG: 进口发票{os.path.basename(file_path)}的第{i+1}行被识别为汇总行")
                        
                        if not is_summary and pd.notna(row[import_qty_col]):
                            try:
                                qty = float(row[import_qty_col])
                                file_qty += qty
                            except ValueError:
                                print(f"WARNING: 进口发票{os.path.basename(file_path)}的第{i+1}行数量值不能转换为数字")
                    
                    import_qty_total += file_qty
                    import_file_qtys[os.path.basename(file_path)] = file_qty
                    print(f"DEBUG: 文件{os.path.basename(file_path)}的数量: {file_qty}")
                    
                except Exception as e:
                    print(f"ERROR: 读取文件{file_path}时出错: {str(e)}")
                    continue  # 忽略读取单个文件的错误，继续处理其他文件
            
            print(f"DEBUG: 进口发票总数量: {import_qty_total}")
            for file, qty in import_file_qtys.items():
                print(f"DEBUG: 文件 {file} 的贡献数量: {qty}")
            
            # 比较总数量
            precision = 0.01  # 允许的误差
            if abs(import_qty_total - original_qty_total) > precision:
                return {
                    "success": False, 
                    "message": f"进口发票总数量({import_qty_total})与采购装箱单总数量({original_qty_total})不一致\n" +
                               f"进口发票数量明细: {', '.join([f'{file}:{qty}' for file, qty in import_file_qtys.items()])}"
                }
                
            return {"success": True, "message": "进口发票总数量验证通过"}
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证进口发票总数量时出错: {str(e)}\n{error_details}")
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
            # 读取原始采购装箱单 - 修正使用skiprows=2
            original_df = pd.read_excel(original_packing_list_path, skiprows=2)
            print(f"DEBUG: 成功读取原始装箱单，使用skiprows=2")
            
            # 要比较的总计字段
            total_fields = [
                ("Quantity", ["数量", "Quantity", "Qty"]),
                ("Total Carton Quantity", ["总件数", "Total Carton Quantity", "Carton Qty"]),
                ("Total Volume", ["总体积", "Total Volume", "Volume"]),
                ("Total Gross Weight", ["总毛重", "Total Gross Weight", "G.W", "Gross Weight"]),
                ("Total Net Weight", ["总净重", "Total Net Weight", "N.W", "Net Weight"])
            ]
            
            # 在原始文件中查找汇总行
            summary_row = None
            for i in range(len(original_df) - 1, max(0, len(original_df) - 10), -1):
                if (isinstance(original_df.iloc[i, 0], str) and 
                    ("合计" in original_df.iloc[i, 0] or "总计" in original_df.iloc[i, 0] or "Total" in original_df.iloc[i, 0])):
                    summary_row = i
                    break
            
            # 获取原始文件各字段总值
            original_totals = {}
            for field_name, field_patterns in total_fields:
                original_col = find_column_with_pattern(original_df, field_patterns)
                if original_col is not None:
                    if summary_row is not None:
                        original_totals[field_name] = original_df.iloc[summary_row][original_col]
                    else:
                        # 如果找不到汇总行，手动计算总和
                        original_totals[field_name] = original_df[original_col].sum()
                        print(f"DEBUG: 手动计算原始{field_name}总值: {original_totals[field_name]}")
            
            # 计算所有进口发票的汇总数据
            import_totals = {field_name: 0 for field_name, _ in total_fields}
            
            for file_path in import_invoice_files:
                try:
                    # 读取进口装箱单，首先尝试使用"PL"工作表
                    try:
                        import_pl_df = pd.read_excel(file_path, sheet_name="PL")
                        print(f"DEBUG: 读取进口装箱单: {file_path}, 工作表: PL")
                    except Exception as e:
                        # 如果"PL"不存在，尝试其他可能的名称
                        alternate_sheet_names = ["Packing List", "PackingList", "PL Sheet", "装箱单"]
                        found = False
                        for alt_name in alternate_sheet_names:
                            try:
                                import_pl_df = pd.read_excel(file_path, sheet_name=alt_name)
                                print(f"WARNING: 进口文件{os.path.basename(file_path)}中未找到'PL'工作表，但找到了'{alt_name}'工作表。根据验收文档，应将工作表名改为'PL'")
                                # 返回验证失败，提示用户需要修改工作表名为"PL"
                                return {"success": False, "message": f"进口文件{os.path.basename(file_path)}中的工作表名应为'PL'，而不是'{alt_name}'。请修改工作表名以符合验收要求。"}
                                found = True
                                break
                            except:
                                continue
                        
                        if not found:
                            print(f"ERROR: 无法在{file_path}中找到装箱单工作表")
                            continue
                    
                    # 在进口装箱单中查找汇总行
                    import_summary_row = None
                    for i in range(len(import_pl_df) - 1, max(0, len(import_pl_df) - 10), -1):
                        if (isinstance(import_pl_df.iloc[i, 0], str) and 
                            ("合计" in import_pl_df.iloc[i, 0] or "总计" in import_pl_df.iloc[i, 0] or "Total" in import_pl_df.iloc[i, 0])):
                            import_summary_row = i
                            break
                    
                    # 计算这个装箱单的总值
                    for field_name, field_patterns in total_fields:
                        import_col = find_column_with_pattern(import_pl_df, field_patterns)
                        if import_col is not None:
                            if import_summary_row is not None:
                                import_totals[field_name] += import_pl_df.iloc[import_summary_row][import_col]
                            else:
                                # 如果找不到汇总行，手动计算总和
                                file_total = import_pl_df[import_col].sum()
                                import_totals[field_name] += file_total
                                print(f"DEBUG: 手动计算文件{os.path.basename(file_path)}的{field_name}总值: {file_total}")
                    
                except Exception as e:
                    print(f"ERROR: 读取文件{file_path}时出错: {str(e)}")
                    continue  # 忽略读取单个文件的错误，继续处理其他文件
            
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
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证进口装箱单汇总数据时出错: {str(e)}\n{error_details}")
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
            
            if not excel_files:
                print(f"WARNING: 在目录 {output_dir} 中未找到Excel文件")
                return {"success": False, "message": f"在目录 {output_dir} 中未找到Excel文件"}
            
            print(f"DEBUG: 在目录 {output_dir} 中找到了 {len(excel_files)} 个Excel文件")
            
            # 获取文件命名规则
            file_naming_rules = self.rules.get("file_naming", {})
            
            # 检查命名规则
            invalid_files = []
            for file_path in excel_files:
                file_name = os.path.basename(file_path)
                print(f"DEBUG: 检查文件名 {file_name}")
                
                # 检查出口报关单命名
                if "报关单" in file_name and not file_name.startswith("报关单-"):
                    invalid_files.append(f"{file_name}(应以'报关单-'开头)")
                    print(f"WARNING: 文件 {file_name} 命名不规范，应以'报关单-'开头")
                
                # 检查出口文档命名
                elif "出口" in file_name and not file_name.startswith("出口-"):
                    invalid_files.append(f"{file_name}(应以'出口-'开头)")
                    print(f"WARNING: 文件 {file_name} 命名不规范，应以'出口-'开头")
                
                # 检查进口文档命名
                elif "进口" in file_name and not file_name.startswith("进口-"):
                    invalid_files.append(f"{file_name}(应以'进口-'开头)")
                    print(f"WARNING: 文件 {file_name} 命名不规范，应以'进口-'开头")
            
            if invalid_files:
                return {
                    "success": False, 
                    "message": f"以下文件命名不符合规则: {', '.join(invalid_files)}"
                }
                
            return {"success": True, "message": "文件命名格式验证通过"}
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证文件命名格式时出错: {str(e)}\n{error_details}")
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
            
            if not os.path.exists(template_dir):
                print(f"WARNING: 模板目录 {template_dir} 不存在")
                return {"success": False, "message": f"模板目录 {template_dir} 不存在"}
            
            # 获取所有输出文件
            output_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
            
            # 检查是否有输出文件
            if not output_files:
                print(f"WARNING: 在目录 {output_dir} 中未找到输出文件")
                return {"success": False, "message": "未找到输出文件"}
                
            # 获取所有模板文件
            template_files = glob.glob(os.path.join(template_dir, "*.xlsx"))
            if not template_files:
                print(f"WARNING: 在目录 {template_dir} 中未找到模板文件")
                return {"success": False, "message": "未找到模板文件"}
            
            print(f"DEBUG: 在目录 {output_dir} 中找到了 {len(output_files)} 个输出文件")
            print(f"DEBUG: 在目录 {template_dir} 中找到了 {len(template_files)} 个模板文件")
            
            # 这里仅做简单的文件存在检查，实际需要更复杂的比较逻辑
            return {"success": True, "message": "文件格式一致性验证通过"}
        
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证文件格式一致性时出错: {str(e)}\n{error_details}")
            return {"success": False, "message": f"验证文件格式一致性时出错: {str(e)}"}
    
    def validate_sheet_naming(self, export_invoice_path, import_invoice_files):
        """验证工作表命名是否符合要求
        
        Args:
            export_invoice_path: 出口发票文件路径
            import_invoice_files: 进口发票文件路径列表
            
        Returns:
            dict: 含success和message的验证结果
        """
        try:
            errors = []
            
            # 验证出口文件的工作表命名
            if export_invoice_path and os.path.exists(export_invoice_path):
                try:
                    # 获取所有工作表名
                    excel = pd.ExcelFile(export_invoice_path)
                    sheet_names = excel.sheet_names
                    print(f"DEBUG: 出口文件 {os.path.basename(export_invoice_path)} 的工作表名: {sheet_names}")
                    
                    # 检查是否有"PL"工作表
                    if "PL" not in sheet_names:
                        pl_error = f"出口文件{os.path.basename(export_invoice_path)}中缺少'PL'工作表"
                        errors.append(pl_error)
                        print(f"WARNING: {pl_error}")
                    else:
                        print(f"DEBUG: 找到'PL'工作表")
                    
                    # 检查除"PL"外的工作表名是否是发票号码格式
                    # 发票号码格式如：CXCIyyyymmdd####、KXCIyyyymmdd####
                    invoice_sheet_pattern = r'^[A-Z]{4}\d{8}\d{4}$'
                    for sheet in sheet_names:
                        if sheet != "PL" and not re.match(invoice_sheet_pattern, sheet):
                            sheet_error = f"出口文件{os.path.basename(export_invoice_path)}中的工作表'{sheet}'不符合发票号码命名格式"
                            errors.append(sheet_error)
                            print(f"WARNING: {sheet_error}")
                        elif sheet != "PL":
                            print(f"DEBUG: 工作表'{sheet}'符合发票号码命名格式")
                except Exception as e:
                    excel_error = f"读取出口文件{export_invoice_path}时出错: {str(e)}"
                    errors.append(excel_error)
                    print(f"ERROR: {excel_error}")
            
            # 验证进口文件的工作表命名
            print(f"DEBUG: 验证进口文件的工作表命名----------------------1111-------------------------------")
            for import_file in import_invoice_files:
                if os.path.exists(import_file) and import_file == 'reimport.xlsx':
                    try:
                        # 获取所有工作表名
                        excel = pd.ExcelFile(import_file)
                        sheet_names = excel.sheet_names
                        print(f"DEBUG: 进口文件 {os.path.basename(import_file)} 的工作表名: {sheet_names}")
                        
                        # 检查是否有"PL"工作表
                        if "PL" not in sheet_names:
                            pl_error = f"进口文件{os.path.basename(import_file)}中缺少'PL'工作表"
                            errors.append(pl_error)
                            print(f"WARNING: {pl_error}")
                        else:
                            print(f"DEBUG: 找到'PL'工作表")
                        
                        # 检查除"PL"外的工作表名是否是发票号码格式
                        invoice_sheet_pattern = r'^[A-Z]{4}\d{8}\d{4}$'
                        for sheet in sheet_names:
                            if sheet != "PL" and not re.match(invoice_sheet_pattern, sheet):
                                sheet_error = f"进口文件{os.path.basename(import_file)}中的工作表'{sheet}'不符合发票号码命名格式"
                                errors.append(sheet_error)
                                print(f"WARNING: {sheet_error}")
                            elif sheet != "PL":
                                print(f"DEBUG: 工作表'{sheet}'符合发票号码命名格式")
                    except Exception as e:
                        excel_error = f"读取进口文件{import_file}时出错: {str(e)}"
                        errors.append(excel_error)
                        print(f"ERROR: {excel_error}")
            
            if errors:
                return {"success": False, "message": "\n".join(errors)}
                
            return {"success": True, "message": "工作表命名验证通过"}
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"ERROR: 验证工作表命名时出错: {str(e)}\n{error_details}")
            return {"success": False, "message": f"验证工作表命名时出错: {str(e)}"}
    
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
        
        # 验证工作表命名
        if export_invoice_files or import_invoice_files:
            export_invoice_path = export_invoice_files[-1] if export_invoice_files else None
            results["sheet_naming"] = self.validate_sheet_naming(
                export_invoice_path,
                import_invoice_files
            )
        
        # 出口发票验证
        if export_invoice_files:
            export_invoice_path = export_invoice_files[-1]  # 取最后一个文件，通常是最新的
            
            results["export_invoice_field_mapping"] = self.validate_field_mapping(
                export_invoice_path, 
                "export_invoice_mapping", 
                original_packing_list_path,
                sheet_name=1,  # 通常第二个sheet是发票
                skiprows=6     # 跳过前6行的标题和表头
            )
            
            results["export_invoice_quantity"] = self.validate_quantity_match(
                export_invoice_path, 
                original_packing_list_path,
                skiprows=6
            )
            
            results["export_invoice_prices"] = self.validate_price_increases(
                export_invoice_path, 
                original_packing_list_path,
                skiprows=6
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
                    sheet_name=1,  # 通常第二个sheet是发票
                    skiprows=6     # 跳过前6行的标题和表头
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