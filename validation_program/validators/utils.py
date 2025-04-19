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
    for col in df.columns:
        if exact:
            if col in patterns:
                return col
        else:
            for pattern in patterns:
                if pattern.lower() in str(col).lower():
                    return col
    return None 