import os
import pandas as pd

def get_output_files(output_dir):
    """获取输出目录中的所有文件"""
    files = []
    for file in os.listdir(output_dir):
        if file.endswith(".xlsx"):
            files.append(os.path.join(output_dir, file))
    return files

def read_excel_to_df(file_path, sheet_name=0):
    """读取Excel文件到DataFrame"""
    return pd.read_excel(file_path, sheet_name=sheet_name)

def compare_numeric_values(value1, value2, precision=0.0001):
    """比较两个数值是否相等(考虑精度)"""
    return abs(value1 - value2) < precision

def find_column_with_pattern(df, patterns, exact=False):
    """查找包含特定模式的列名
    
    Args:
        df: DataFrame
        patterns: 模式列表
        exact: 是否精确匹配
        
    Returns:
        匹配的列名或None
    """
    try:
        # 检查是否有多层表头
        has_multiindex = isinstance(df.columns, pd.MultiIndex)
        print(f"DEBUG: 查找模式 {patterns}, 表格是否有多层表头: {has_multiindex}")
        
        # 首先尝试使用精确匹配
        for col in df.columns:
            # 多层表头情况
            if has_multiindex:
                # 检查元组中的每个部分
                for pattern in patterns:
                    if pattern in col:  # 如果模式是元组中的一部分
                        print(f"DEBUG: 精确匹配到列: {col} with pattern {pattern}")
                        return col
            else:
                # 单层表头的情况
                if col in patterns:
                    print(f"DEBUG: 精确匹配到列: {col}")
                    return col
        
        # 如果未找到精确匹配，则尝试模糊匹配
        if not exact:
            for col in df.columns:
                # 多层表头的情况
                if has_multiindex:
                    # 将元组转换为字符串进行匹配
                    col_str = ' '.join([str(c).lower() for c in col])
                    for pattern in patterns:
                        pattern_str = str(pattern).lower()
                        if pattern_str in col_str:
                            print(f"DEBUG: 模糊匹配到列(多层表头): {col} with pattern {pattern}")
                            return col
                else:
                    # 单层表头的情况
                    col_str = str(col).lower()
                    for pattern in patterns:
                        pattern_str = str(pattern).lower()
                        if pattern_str in col_str:
                            print(f"DEBUG: 模糊匹配到列(单层表头): {col} with pattern {pattern}")
                            return col
        
        # 特殊处理常见模式的变体
        for col in df.columns:
            # 多层表头的情况
            if has_multiindex:
                col_str = ' '.join([str(c).lower() for c in col])
            else:
                col_str = str(col).lower()
                
            # 净重相关的特殊处理
            net_weight_search = any(p.lower() in ['net weight', 'n.w', 'n/w', '净重'] for p in patterns)
            if net_weight_search:
                if ('net' in col_str and 'weight' in col_str) or 'n.w' in col_str or 'n/w' in col_str or '净重' in col_str:
                    print(f"DEBUG: 特殊匹配到净重列: {col}")
                    return col
                    
            # 运费相关的特殊处理
            freight_search = any(p.lower() in ['freight', 'total freight', '运费', '总运费'] for p in patterns)
            if freight_search:
                if 'freight' in col_str or '运费' in col_str:
                    print(f"DEBUG: 特殊匹配到运费列: {col}")
                    return col
                    
        print(f"DEBUG: 未找到匹配列，可用列名: {list(df.columns)}")
        return None
    except Exception as e:
        print(f"ERROR: 查找列时出错: {str(e)}")
        raise Exception(f"查找模式{patterns}的列时出错: {str(e)}") 