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
        
        # 特殊处理 factory 和 project 列的搜索
        factory_related_patterns = ["Plant Location", "factory", "工厂", "工厂地点", "daman/silvass", "送达方", "目的地", "Location", "plant", "厂区"]
        project_related_patterns = ["Project", "project", "项目", "项目名称", "所属项目", "program", "program name", "计划名称"]
        
        if any(p.lower() in [pat.lower() for pat in patterns] for p in factory_related_patterns):
            print("DEBUG: 正在搜索工厂相关列")
            extended_patterns = factory_related_patterns
            # 优先检查是否有明确的"factory"列
            if "factory" in df.columns:
                print(f"DEBUG: 找到明确的工厂列 'factory'")
                return "factory"
        elif any(p.lower() in [pat.lower() for pat in patterns] for p in project_related_patterns):
            print("DEBUG: 正在搜索项目相关列")
            extended_patterns = project_related_patterns
            # 优先检查是否有明确的"project"列
            if "project" in df.columns:
                print(f"DEBUG: 找到明确的项目列 'project'")
                return "project"
        else:
            extended_patterns = patterns
        
        # 首先尝试使用精确匹配
        for col in df.columns:
            # 多层表头情况
            if has_multiindex:
                # 检查元组中的每个部分
                for pattern in extended_patterns:
                    if pattern in col:  # 如果模式是元组中的一部分
                        print(f"DEBUG: 精确匹配到列: {col} with pattern {pattern}")
                        return col
            else:
                # 单层表头的情况
                if col in extended_patterns:
                    print(f"DEBUG: 精确匹配到列: {col}")
                    return col
        
        # 如果未找到精确匹配，则尝试模糊匹配
        if not exact:
            for col in df.columns:
                # 多层表头的情况
                if has_multiindex:
                    # 将元组转换为字符串进行匹配
                    col_str = ' '.join([str(c).lower() for c in col])
                    for pattern in extended_patterns:
                        pattern_str = str(pattern).lower()
                        if pattern_str in col_str:
                            print(f"DEBUG: 模糊匹配到列(多层表头): {col} with pattern {pattern}")
                            return col
                else:
                    # 单层表头的情况
                    col_str = str(col).lower()
                    for pattern in extended_patterns:
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
                    
            # 工厂相关的特殊处理
            factory_search = any(p.lower() in ['factory', 'plant location', '工厂', '工厂地点'] for p in patterns)
            if factory_search:
                if ('factory' in col_str or 'plant' in col_str or '工厂' in col_str or '厂区' in col_str or 
                    'location' in col_str or '地点' in col_str or 'daman' in col_str or 'silvass' in col_str):
                    print(f"DEBUG: 特殊匹配到工厂列: {col}")
                    return col
                    
            # 项目相关的特殊处理
            project_search = any(p.lower() in ['project', '项目', '项目名称'] for p in patterns)
            if project_search:
                if 'project' in col_str or 'program' in col_str or '项目' in col_str or '计划' in col_str:
                    print(f"DEBUG: 特殊匹配到项目列: {col}")
                    return col
                    
        print(f"DEBUG: 未找到匹配列，可用列名: {list(df.columns)}")
        return None
    except Exception as e:
        print(f"ERROR: 查找列时出错: {str(e)}")
        raise Exception(f"查找模式{patterns}的列时出错: {str(e)}") 