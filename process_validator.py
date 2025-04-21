import pandas as pd
import os
import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

# 支持比较数值
def compare_numeric_values(value1, value2, precision=0.0001):
    """比较两个数值是否相等(考虑精度)"""
    return abs(value1 - value2) < precision

# 查找列
def find_column_with_pattern(df, patterns):
    """查找包含特定模式的列名"""
    # 首先尝试精确匹配
    for pattern in patterns:
        if pattern in df.columns:
            return pattern
            
    # 如果精确匹配失败，尝试模糊匹配
    for pattern in patterns:
        for col in df.columns:
            if pattern.lower() in str(col).lower():
                return col
                
    return None

def validate_factory_split(self, cif_invoice_path, import_invoice_dir):
    """校验工厂拆分是否正确
    
    Args:
        cif_invoice_path: CIF发票路径
        import_invoice_dir: 进口发票目录
        
    Returns:
        tuple: (bool, str) 是否通过校验，错误信息
    """
    try:
        # 读取CIF发票，查找工厂列
        logging.info(f"读取CIF发票: {cif_invoice_path}")
        cif_df = pd.read_excel(cif_invoice_path)
        logging.info(f"CIF发票列名: {list(cif_df.columns)}")
        
        # 尝试识别工厂列
        factory_col = None
        # 尝试常用的工厂列名模式
        factory_patterns = ["Plant Location", "工厂地点", "工厂", "factory", "Factory", "FACTORY", 
                            "Plant", "plant", "Location", "location", "工厂名称", "Supplier"]
        
        for pattern in factory_patterns:
            for col in cif_df.columns:
                if pattern in str(col):
                    factory_col = col
                    logging.info(f"找到工厂列: {col}")
                    break
            if factory_col:
                break
                
        # 如果直接匹配没找到，尝试更广泛的匹配
        if not factory_col:
            for col in cif_df.columns:
                col_str = str(col).lower()
                if 'factory' in col_str or '工厂' in col_str or 'plant' in col_str or 'location' in col_str or 'supplier' in col_str:
                    factory_col = col
                    logging.info(f"通过部分匹配找到工厂列: {col}")
                    break
        
        # 检查进口发票目录是否存在
        if not os.path.exists(import_invoice_dir):
            return False, f"进口发票目录不存在: {import_invoice_dir}"
            
        # 获取目录下所有Excel文件
        import_files = [f for f in os.listdir(import_invoice_dir) 
                        if f.endswith('.xlsx') or f.endswith('.xls')]
        
        if not import_files:
            return False, f"进口发票目录中没有找到Excel文件: {import_invoice_dir}"
            
        logging.info(f"进口发票目录中的文件: {import_files}")
        
        # 检查是否有工厂拆分的标志
        has_factory_split = False
        for f in import_files:
            if 'reimport_' in f:
                has_factory_split = True
                break
        
        # 特殊情况：如果只有一个默认工厂文件，认为是合法的
        if len(import_files) == 1 and ('默认工厂' in import_files[0] or '默认工厂'.encode('utf-8').decode('latin1') in import_files[0]):
            return True, "使用了默认工厂，验证通过"
            
        # 如果没有工厂列但有多个进口文件，说明可能是手动拆分的
        if not factory_col and len(import_files) > 1:
            logging.warning("未找到工厂列，但有多个进口文件，可能是手动拆分")
            return True, "未找到工厂列但检测到多个进口文件，验证通过（可能是手动拆分）"
        
        # 如果找到工厂列，验证每个工厂是否都有对应的文件
        if factory_col:
            factories = cif_df[factory_col].dropna().unique()
            logging.info(f"CIF发票中的工厂: {factories}")
            
            # 如果没有工厂值但使用了默认工厂文件，认为是合法的
            if len(factories) == 0:
                for f in import_files:
                    if '默认工厂' in f or '默认工厂'.encode('utf-8').decode('latin1') in f:
                        return True, "CIF中无工厂值，使用了默认工厂，验证通过"
            
            # 验证每个工厂是否都有对应的文件
            missing_factories = []
            for factory in factories:
                factory_str = str(factory).replace('/', '_').replace('\\', '_')
                found = False
                
                for f in import_files:
                    # 处理可能的编码问题
                    try:
                        if factory_str in f:
                            found = True
                            break
                    except:
                        # 尝试不同的编码处理
                        if factory_str.encode('utf-8').decode('latin1') in f:
                            found = True
                            break
                
                if not found:
                    missing_factories.append(factory)
            
            if missing_factories:
                return False, f"以下工厂没有对应的进口发票文件: {missing_factories}"
            
            return True, "所有工厂都有对应的进口发票文件，验证通过"
        
        # 如果找不到工厂列但使用了默认工厂，也认为是合法的
        for f in import_files:
            if '默认工厂' in f or '默认工厂'.encode('utf-8').decode('latin1') in f:
                return True, "未找到工厂列，使用了默认工厂，验证通过"
        
        return False, "未找到工厂列，且没有默认工厂的进口发票文件"
        
    except Exception as e:
        logging.error(f"验证工厂拆分时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, f"验证失败: {str(e)}"


def validate_cif_price_calculation(self, cif_invoice_path):
    """验证CIF价格计算
    
    Args:
        cif_invoice_path: CIF发票文件路径
        
    Returns:
        tuple: (bool, str) 是否通过校验，错误信息
    """
    try:
        # 读取CIF发票
        cif_df = pd.read_excel(cif_invoice_path)
        
        # 找到FOB单价、单个物料保险费、单个物料运费、CIF单价列
        # 修改列名模式以匹配实际文件中的列名
        fob_price_col = find_column_with_pattern(cif_df, ["FOB Unit Price", "FOB单价", "FOB总价"])
        insurance_freight_col = find_column_with_pattern(cif_df, ["Insurance", "该项对应的运保费", "运保费"])
        cif_price_col = find_column_with_pattern(cif_df, ["CIF Unit Price", "CIF单价", "CIF总价(FOB总价+运保费)"])
        
        if None in [fob_price_col, insurance_freight_col, cif_price_col]:
            logging.warning(f"未找到所有需要的价格列, 列名: {list(cif_df.columns)}")
            return False, "未找到所有需要的价格列，无法验证CIF价格计算"
        
        # 验证每行CIF单价是否等于FOB单价+单个物料保险费+单个物料运费
        invalid_rows = []
        for idx, row in cif_df.iterrows():
            if pd.isna(row[fob_price_col]) or pd.isna(row[insurance_freight_col]) or pd.isna(row[cif_price_col]):
                continue
            
            expected_cif = row[fob_price_col] + row[insurance_freight_col]
            actual_cif = row[cif_price_col]
            
            # 允许小误差
            if not compare_numeric_values(expected_cif, actual_cif, 0.01):
                invalid_rows.append(idx + 1)  # +1是因为0基索引
        
        if invalid_rows:
            return False, f"以下行的CIF价格计算不正确: {', '.join(map(str, invalid_rows))}"
            
        return True, "CIF价格计算验证通过"
    except Exception as e:
        logging.error(f"验证CIF价格计算时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, f"验证失败: {str(e)}"


def validate_merge_logic(self, cif_invoice_path, export_invoice_path):
    """验证相同物料编号和价格的合并逻辑
    
    Args:
        cif_invoice_path: CIF发票文件路径
        export_invoice_path: 出口发票文件路径
        
    Returns:
        tuple: (bool, str) 是否通过校验，错误信息
    """
    try:
        # 读取CIF发票
        cif_df = pd.read_excel(cif_invoice_path)
        
        # 读取出口发票 - 跳过前6行，因为这是发票头部信息
        try:
            export_df = pd.read_excel(export_invoice_path, sheet_name=1, skiprows=6)
            logging.info(f"成功读取出口发票第2个Sheet，共{len(export_df)}行")
        except Exception as e:
            # 如果指定sheet读取失败，尝试第一个sheet
            logging.warning(f"读取出口发票第2个Sheet失败，尝试第1个Sheet: {str(e)}")
            export_df = pd.read_excel(export_invoice_path, skiprows=6)
            logging.info(f"成功读取出口发票第1个Sheet，共{len(export_df)}行")
        
        # 找到物料编号列
        cif_part_col = find_column_with_pattern(cif_df, ["Material code", "物料编号", "料号"])
        export_part_col = find_column_with_pattern(export_df, ["Material code", "物料编号", "料号"])
        
        # 找到单价列 - 修改列名匹配模式
        # 对于CIF文件，使用Unit Price而不是单价USD数值
        cif_price_col = find_column_with_pattern(cif_df, ["Unit Price", "单价", "采购单价"])
        export_price_col = find_column_with_pattern(export_df, ["Unit Price", "单价"])
        
        # 找到数量列
        cif_qty_col = find_column_with_pattern(cif_df, ["Qty", "数量"])
        export_qty_col = find_column_with_pattern(export_df, ["Qty", "数量"])
        
        if None in [cif_part_col, export_part_col, cif_price_col, export_price_col, cif_qty_col, export_qty_col]:
            missing_cols = []
            if cif_part_col is None: missing_cols.append("CIF发票物料编号列")
            if export_part_col is None: missing_cols.append("出口发票物料编号列")
            if cif_price_col is None: missing_cols.append("CIF发票单价列")
            if export_price_col is None: missing_cols.append("出口发票单价列")
            if cif_qty_col is None: missing_cols.append("CIF发票数量列")
            if export_qty_col is None: missing_cols.append("出口发票数量列")
            
            logging.warning(f"未找到所有需要的列: {missing_cols}")
            logging.warning(f"CIF发票列名: {list(cif_df.columns)}")
            logging.warning(f"出口发票列名: {list(export_df.columns)}")
            
            return False, f"未找到所有需要的列，无法验证合并逻辑: {', '.join(missing_cols)}"
        
        logging.info(f"CIF文件使用列: 物料={cif_part_col}, 单价={cif_price_col}, 数量={cif_qty_col}")
        logging.info(f"出口文件使用列: 物料={export_part_col}, 单价={export_price_col}, 数量={export_qty_col}")
        
        # 统计CIF发票中相同物料编号和单价的数量总和
        cif_grouped = cif_df.groupby([cif_part_col, cif_price_col])[cif_qty_col].sum().reset_index()
        
        # 检查出口发票中每个物料
        for _, export_row in export_df.iterrows():
            export_part = export_row[export_part_col]
            export_price = export_row[export_price_col]
            export_qty = export_row[export_qty_col]
            
            # 跳过空行
            if pd.isna(export_part) or pd.isna(export_price) or pd.isna(export_qty):
                continue
            
            # 在CIF分组中查找对应的物料和价格
            found = False
            for _, cif_row in cif_grouped.iterrows():
                cif_part = cif_row[cif_part_col]
                cif_price = cif_row[cif_price_col]
                cif_qty = cif_row[cif_qty_col]
                
                # 价格可能有精度差异，使用近似比较
                if cif_part == export_part and compare_numeric_values(cif_price, export_price, 0.0001):
                    found = True
                    
                    # 检查数量是否匹配
                    if not compare_numeric_values(cif_qty, export_qty, 0.01):
                        return False, f"物料{export_part}的合并数量不正确: CIF总计({cif_qty}) vs 出口({export_qty})"
                    
                    break
            
            if not found:
                # 输出更多debug信息，帮助诊断问题
                logging.warning(f"未找到物料匹配: {export_part}, 单价: {export_price}")
                # 查找是否有相同物料但价格不同的情况
                matching_parts = cif_df[cif_df[cif_part_col] == export_part]
                if not matching_parts.empty:
                    logging.warning(f"找到相同物料但价格不同: {matching_parts[[cif_part_col, cif_price_col]].values.tolist()}")
                
                return False, f"未找到物料{export_part}单价{export_price}在CIF发票中的对应项"
        
        return True, "物料合并逻辑验证通过"
    except Exception as e:
        logging.error(f"验证物料合并逻辑时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, f"验证失败: {str(e)}" 