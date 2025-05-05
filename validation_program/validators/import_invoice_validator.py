import pandas as pd
from .utils import find_column_with_pattern, compare_numeric_values

class ImportInvoiceValidator:
    """进口发票验证器"""

    def __init__(self):
        """初始化验证器"""
        pass

    def validate_row_merging(self, import_invoice_path):
        """验证相同Part Number和Unit Price的行是否已合并

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到Part Number和Unit Price列
            part_number_col = find_column_with_pattern(import_df, ["Part Number", "料号"])
            unit_price_col = find_column_with_pattern(import_df, ["Unit Price (CIF, USD)", "单价"])

            if part_number_col is None or unit_price_col is None:
                return {"success": False, "message": "未找到Part Number或Unit Price列，无法验证行合并"}

            # 检查是否有重复的Part Number和Unit Price组合
            # 先将Unit Price四舍五入到4位小数，以确保比较的一致性
            if unit_price_col in import_df.columns:
                import_df['Rounded_Price'] = import_df[unit_price_col].round(4)
            else:
                return {"success": False, "message": f"Unit Price列 '{unit_price_col}' 不在数据框中"}

            # 过滤掉NaN值和总计行
            valid_rows = import_df.dropna(subset=[part_number_col, 'Rounded_Price'])
            valid_rows = valid_rows[~valid_rows[part_number_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD', na=False, regex=True)]

            # 检查是否有重复的Part Number和Rounded_Price组合
            duplicates = valid_rows.duplicated(subset=[part_number_col, 'Rounded_Price'], keep=False)

            if duplicates.any():
                duplicate_rows = valid_rows[duplicates]
                duplicate_items = []

                for part, price in duplicate_rows[[part_number_col, 'Rounded_Price']].drop_duplicates().values:
                    duplicate_items.append(f"Part Number: {part}, Unit Price: {price}")

                return {
                    "success": False,
                    "message": f"发现未合并的相同Part Number和Unit Price组合:\n" + "\n".join(duplicate_items)
                }

            return {"success": True, "message": "所有相同Part Number和Unit Price的行已正确合并"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证行合并时出错: {str(e)}, 行号: {error_line}"}

    def validate_sn_numbering(self, import_invoice_path):
        """验证S/N是否从1开始编号

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到S/N列
            sn_col = find_column_with_pattern(import_df, ["S/N", "序号"])

            if sn_col is None:
                return {"success": False, "message": "未找到S/N列，无法验证编号"}

            # 获取数据行（排除汇总行和页脚行）
            data_rows = import_df[~import_df[sn_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD|PACKED IN|NET WEIGHT|GROSS WEIGHT', na=False, regex=True)].copy()

            if data_rows.empty:
                return {"success": False, "message": "未找到有效的数据行，无法验证S/N编号"}

            # 检查S/N是否从1开始，并且是连续的
            try:
                # 尝试将S/N转换为数字
                sn_values = pd.to_numeric(data_rows[sn_col], errors='coerce')

                # 过滤掉非数字的值
                sn_values = sn_values.dropna()

                if sn_values.empty:
                    return {"success": False, "message": "S/N列不包含有效的数字，无法验证编号"}

                # 检查是否从1开始
                if sn_values.min() != 1:
                    return {"success": False, "message": f"S/N编号不是从1开始，最小值为: {sn_values.min()}"}

                # 检查是否连续
                expected_values = list(range(1, len(sn_values) + 1))
                actual_values = sorted(sn_values.tolist())

                if expected_values != actual_values:
                    missing_values = set(expected_values) - set(actual_values)
                    extra_values = set(actual_values) - set(expected_values)

                    message = "S/N编号不连续。"
                    if missing_values:
                        message += f"缺少的编号: {sorted(missing_values)}. "
                    if extra_values:
                        message += f"多余的编号: {sorted(extra_values)}."

                    return {"success": False, "message": message}

                return {"success": True, "message": f"S/N编号正确，从1开始连续编号到{len(sn_values)}"}

            except Exception as e:
                return {"success": False, "message": f"验证S/N编号时出错: {str(e)}"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证S/N编号时出错: {str(e)}, 行号: {error_line}"}

    def validate_description_field(self, import_invoice_path, original_packing_list_path):
        """验证进口发票是否使用了进口清关货描作为描述字段

        Args:
            import_invoice_path: 进口发票文件路径
            original_packing_list_path: 原始装箱单文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到描述列
            desc_col = find_column_with_pattern(import_df, ["Commodity Description (Customs)", "Description", "描述"])
            part_number_col = find_column_with_pattern(import_df, ["Part Number", "料号"])

            if desc_col is None:
                return {"success": False, "message": "未找到描述列，无法验证描述字段"}

            if part_number_col is None:
                return {"success": False, "message": "未找到Part Number列，无法验证描述字段"}

            # 读取原始装箱单
            try:
                # 尝试使用多级表头读取
                original_df = pd.read_excel(original_packing_list_path, header=[1,2], skiprows=[0])
                print(f"DEBUG: 使用多级表头读取原始装箱单成功")
            except Exception as e:
                print(f"DEBUG: 使用多级表头读取失败，尝试使用skiprows=2: {str(e)}")
                original_df = pd.read_excel(original_packing_list_path, skiprows=2)
                print(f"DEBUG: 使用skiprows=2读取原始装箱单成功")

            # 找到原始装箱单中的进口清关货描列和料号列
            customs_desc_col = None
            for col in original_df.columns:
                col_str = str(col).lower()
                if '进口清关货描' in col_str or 'commodity description (customs)' in col_str:
                    customs_desc_col = col
                    break

            original_part_col = find_column_with_pattern(original_df, ["Part Number", "料号", "Material code"])

            if customs_desc_col is None:
                return {"success": False, "message": "原始装箱单中未找到进口清关货描列，无法验证描述字段"}

            if original_part_col is None:
                return {"success": False, "message": "原始装箱单中未找到料号列，无法验证描述字段"}

            # 获取进口发票中的数据行（排除汇总行和页脚行）
            data_rows = import_df[~import_df[part_number_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD', na=False, regex=True)].copy()

            if data_rows.empty:
                return {"success": False, "message": "未找到有效的数据行，无法验证描述字段"}

            # 检查每个Part Number的描述是否与原始装箱单中的进口清关货描匹配
            mismatched_items = []

            for _, row in data_rows.iterrows():
                try:
                    part_number = row[part_number_col]
                    description = row[desc_col]

                    # 跳过空值
                    if pd.isna(part_number) or pd.isna(description):
                        continue

                    # 在原始装箱单中查找对应的料号
                    matching_rows = original_df[original_df[original_part_col] == part_number]

                    if matching_rows.empty:
                        mismatched_items.append(f"Part Number: {part_number} - 在原始装箱单中未找到")
                        continue

                    # 获取原始装箱单中的进口清关货描
                    original_desc = matching_rows.iloc[0][customs_desc_col]

                    # 如果原始描述为空，跳过此项
                    if pd.isna(original_desc) or str(original_desc).strip() == '':
                        continue

                    # 比较描述是否匹配
                    if str(description).strip() != str(original_desc).strip():
                        mismatched_items.append(f"Part Number: {part_number} - 描述不匹配。进口发票: '{description}', 原始装箱单进口清关货描: '{original_desc}'")
                except Exception as e:
                    print(f"DEBUG: 处理行时出错: {str(e)}, 行内容: {row}")

            if mismatched_items:
                return {
                    "success": False,
                    "message": f"以下物料的描述与原始装箱单中的进口清关货描不匹配:\n" + "\n".join(mismatched_items)
                }

            return {"success": True, "message": "所有物料的描述与原始装箱单中的进口清关货描匹配"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证描述字段时出错: {str(e)}, 行号: {error_line}"}

    def validate_quantity_sum(self, import_invoice_path):
        """验证合并行后的数量是否正确求和

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到数量列和总计行
            qty_col = find_column_with_pattern(import_df, ["Quantity", "数量"])
            part_number_col = find_column_with_pattern(import_df, ["Part Number", "料号"])
            desc_col = find_column_with_pattern(import_df, ["Commodity Description (Customs)", "Description", "描述"])

            if qty_col is None:
                return {"success": False, "message": "未找到数量列，无法验证数量求和"}

            if part_number_col is None or desc_col is None:
                return {"success": False, "message": "未找到Part Number或描述列，无法验证数量求和"}

            # 找到总计行
            total_row = None
            for idx, row in import_df.iterrows():
                if pd.notna(row[desc_col]) and str(row[desc_col]).strip() in ['Total', '合计']:
                    total_row = row
                    break

            if total_row is None:
                return {"success": False, "message": "未找到总计行，无法验证数量求和"}

            # 获取数据行（排除总计行和页脚行）
            data_rows = import_df[~import_df[desc_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD|PACKED IN|NET WEIGHT|GROSS WEIGHT', na=False, regex=True)].copy()

            if data_rows.empty:
                return {"success": False, "message": "未找到有效的数据行，无法验证数量求和"}

            # 计算数据行的数量总和
            try:
                data_qty_sum = pd.to_numeric(data_rows[qty_col], errors='coerce').sum()
                total_qty = pd.to_numeric(total_row[qty_col], errors='coerce')

                # 比较总和是否匹配（允许0.01的误差）
                if not compare_numeric_values(data_qty_sum, total_qty, 0.01):
                    return {
                        "success": False,
                        "message": f"数量总和不匹配。数据行总和: {data_qty_sum}, 总计行: {total_qty}"
                    }

                return {"success": True, "message": f"数量总和验证通过，总数量: {total_qty}"}

            except Exception as e:
                return {"success": False, "message": f"计算数量总和时出错: {str(e)}"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证数量求和时出错: {str(e)}, 行号: {error_line}"}

    def validate_amount_sum(self, import_invoice_path):
        """验证合并行后的金额是否正确求和

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到金额列和总计行
            amount_col = find_column_with_pattern(import_df, ["Total Amount (CIF, USD)", "金额", "总金额"])
            desc_col = find_column_with_pattern(import_df, ["Commodity Description (Customs)", "Description", "描述"])

            if amount_col is None:
                return {"success": False, "message": "未找到金额列，无法验证金额求和"}

            if desc_col is None:
                return {"success": False, "message": "未找到描述列，无法验证金额求和"}

            # 找到总计行
            total_row = None
            for idx, row in import_df.iterrows():
                if pd.notna(row[desc_col]) and str(row[desc_col]).strip() in ['Total', '合计']:
                    total_row = row
                    break

            if total_row is None:
                return {"success": False, "message": "未找到总计行，无法验证金额求和"}

            # 获取数据行（排除总计行和页脚行）
            data_rows = import_df[~import_df[desc_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD|PACKED IN|NET WEIGHT|GROSS WEIGHT', na=False, regex=True)].copy()

            if data_rows.empty:
                return {"success": False, "message": "未找到有效的数据行，无法验证金额求和"}

            # 计算数据行的金额总和
            try:
                data_amount_sum = pd.to_numeric(data_rows[amount_col], errors='coerce').sum()
                total_amount = pd.to_numeric(total_row[amount_col], errors='coerce')

                # 比较总和是否匹配（允许0.01的误差）
                if not compare_numeric_values(data_amount_sum, total_amount, 0.01):
                    return {
                        "success": False,
                        "message": f"金额总和不匹配。数据行总和: {data_amount_sum}, 总计行: {total_amount}"
                    }

                return {"success": True, "message": f"金额总和验证通过，总金额: {total_amount}"}

            except Exception as e:
                return {"success": False, "message": f"计算金额总和时出错: {str(e)}"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证金额求和时出错: {str(e)}, 行号: {error_line}"}

    def validate_net_weight_sum(self, import_invoice_path):
        """验证合并行后的净重是否正确求和

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            dict: 含success和message的验证结果
        """
        try:
            # 读取进口发票
            import_df = self.read_import_invoice(import_invoice_path)

            if import_df is None:
                return {"success": False, "message": "无法读取进口发票，请检查文件格式"}

            # 找到净重列和总计行
            net_weight_col = find_column_with_pattern(import_df, ["Total Net Weight (kg)", "净重"])
            desc_col = find_column_with_pattern(import_df, ["Commodity Description (Customs)", "Description", "描述"])

            if net_weight_col is None:
                return {"success": False, "message": "未找到净重列，无法验证净重求和"}

            if desc_col is None:
                return {"success": False, "message": "未找到描述列，无法验证净重求和"}

            # 找到总计行
            total_row = None
            for idx, row in import_df.iterrows():
                if pd.notna(row[desc_col]) and str(row[desc_col]).strip() in ['Total', '合计']:
                    total_row = row
                    break

            if total_row is None:
                return {"success": False, "message": "未找到总计行，无法验证净重求和"}

            # 获取数据行（排除总计行和页脚行）
            data_rows = import_df[~import_df[desc_col].astype(str).str.contains('Total|合计|Amount in Words|SAY USD|PACKED IN|NET WEIGHT|GROSS WEIGHT', na=False, regex=True)].copy()

            if data_rows.empty:
                return {"success": False, "message": "未找到有效的数据行，无法验证净重求和"}

            # 计算数据行的净重总和
            try:
                data_net_weight_sum = pd.to_numeric(data_rows[net_weight_col], errors='coerce').sum()
                total_net_weight = pd.to_numeric(total_row[net_weight_col], errors='coerce')

                # 比较总和是否匹配（允许0.01的误差）
                if not compare_numeric_values(data_net_weight_sum, total_net_weight, 0.01):
                    return {
                        "success": False,
                        "message": f"净重总和不匹配。数据行总和: {data_net_weight_sum}, 总计行: {total_net_weight}"
                    }

                return {"success": True, "message": f"净重总和验证通过，总净重: {total_net_weight}"}

            except Exception as e:
                return {"success": False, "message": f"计算净重总和时出错: {str(e)}"}

        except Exception as e:
            import traceback
            error_line = traceback.extract_tb(e.__traceback__)[-1][1]
            return {"success": False, "message": f"验证净重求和时出错: {str(e)}, 行号: {error_line}"}

    def read_import_invoice(self, import_invoice_path):
        """读取进口发票

        Args:
            import_invoice_path: 进口发票文件路径

        Returns:
            DataFrame: 进口发票数据框或None
        """
        import_df = None
        # 尝试不同的skiprows值读取进口发票
        for sheet_idx in range(1, 5):  # 尝试不同的工作表
            try:
                sheet_names = pd.ExcelFile(import_invoice_path).sheet_names
                if len(sheet_names) <= sheet_idx:
                    continue
                sheet_name = sheet_names[sheet_idx]
                for skiprows in range(0, 15):  # 尝试不同的跳过行数
                    try:
                        temp_df = pd.read_excel(import_invoice_path, sheet_name=sheet_name, skiprows=skiprows)
                        # 检查是否包含关键列名
                        if any('S/N' in str(col) for col in temp_df.columns) or any('Part Number' in str(col) for col in temp_df.iloc[0].values if pd.notna(col)):
                            # 如果第一行包含列名，则重新读取
                            if any('S/N' in str(col) for col in temp_df.iloc[0].values if pd.notna(col)):
                                import_df = pd.read_excel(import_invoice_path, sheet_name=sheet_name, skiprows=skiprows+1, header=0)
                                # 重命名列
                                for i, col in enumerate(import_df.columns):
                                    if i < len(temp_df.columns) and pd.notna(temp_df.iloc[0, i]) and str(temp_df.iloc[0, i]).strip():
                                        import_df.rename(columns={col: str(temp_df.iloc[0, i]).strip()}, inplace=True)
                            else:
                                import_df = temp_df

                            print(f"DEBUG: 成功读取进口发票 {import_invoice_path}, 工作表 {sheet_name}, 跳过行数 {skiprows}")
                            print(f"DEBUG: 进口发票列名: {import_df.columns.tolist()}")
                            print(f"DEBUG: 进口发票前5行:\n{import_df.head()}")
                            return import_df
                    except Exception as e:
                        continue
            except Exception as e:
                continue

        return None

    def validate_all(self, import_invoice_path, original_packing_list_path):
        """运行所有进口发票验证

        Args:
            import_invoice_path: 进口发票文件路径
            original_packing_list_path: 原始装箱单文件路径

        Returns:
            dict: 包含所有验证结果的字典
        """
        results = {}

        # 验证行合并
        results["import_row_merging"] = self.validate_row_merging(import_invoice_path)

        # 验证S/N编号
        results["import_sn_numbering"] = self.validate_sn_numbering(import_invoice_path)

        # 验证描述字段
        results["import_description_field"] = self.validate_description_field(import_invoice_path, original_packing_list_path)

        # 验证数量求和
        results["import_quantity_sum"] = self.validate_quantity_sum(import_invoice_path)

        # 验证金额求和
        results["import_amount_sum"] = self.validate_amount_sum(import_invoice_path)

        # 验证净重求和
        results["import_net_weight_sum"] = self.validate_net_weight_sum(import_invoice_path)

        return results
