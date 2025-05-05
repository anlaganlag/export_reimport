import os
import pandas as pd

def get_output_files(output_dir):
    """获取输出目录中的所有文件"""
    files = []
    for file in os.listdir(output_dir):
        if file.endswith(".xlsx"):
            files.append(os.path.join(output_dir, file))
    return files

def read_excel_to_df(file_path, sheet_name=0, skiprows=None):
    """读取Excel文件到DataFrame
    
    Args:
        file_path: Excel文件路径
        sheet_name: 工作表名称或索引，默认为0
        skiprows: 要跳过的行数，默认为None
    
    Returns:
        pandas.DataFrame: 读取的数据表
    """
    return pd.read_excel(file_path, sheet_name=sheet_name, skiprows=skiprows)

def compare_numeric_values(value1, value2, precision=0.0001):
    """比较两个数值是否相等(考虑精度)"""
    return abs(value1 - value2) < precision

def find_column_with_pattern(df, patterns):
    """在DataFrame中查找包含指定模式的列

    Args:
        df: pandas DataFrame
        patterns: 要匹配的模式列表

    Returns:
        匹配的列名，如果找不到则返回None
    """
    if df is None or patterns is None:
        return None
        
    # 先尝试精确匹配
    for pattern in patterns:
        if pattern in df.columns:
            return pattern
    
    # 如果精确匹配失败，尝试部分匹配
    for pattern in patterns:
        for col in df.columns:
            if pattern in col or col in pattern:
                return col
                
    # 如果都失败了，尝试不区分大小写的匹配
    pattern_lower = [p.lower() for p in patterns]
    col_lower = [c.lower() for c in df.columns]
    
    for p in pattern_lower:
        for i, c in enumerate(col_lower):
            if p in c or c in p:
                return df.columns[i]
    
    return None

def find_value_by_fieldname(df, fieldnames, field_col=0, value_col=1):
    """
    在DataFrame中通过字段名查找对应的值（模糊匹配，A列内容包含任一字段名即可）
    :param df: DataFrame
    :param fieldnames: 字段名列表
    :param field_col: 字段名所在列索引，默认0
    :param value_col: 值所在列索引，默认1
    :return: 找到的值或None
    """
    for i, v in enumerate(df.iloc[:, field_col]):
        if pd.notna(v):
            v_str = str(v).strip()
            for name in fieldnames:
                if name in v_str:
                    return df.iloc[i, value_col]
    return None