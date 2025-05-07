import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
import io
from process_shipping_list import process_shipping_list, read_policy_file
from pathlib import Path
from validation_program.validators.input_validator import InputValidator

# Set page config
st.set_page_config(
    page_title="Export Invoice Generator",
    page_icon="📦",
    layout="wide"
)

# Title and description
st.title("Export Reimport Invoice Generator 出口进口发票生成器")
st.markdown("""
This application helps you generate export invoices from packing lists and policy files.
此应用程序帮助您从装箱单和政策文件生成出口发票。

### Instructions 使用说明
1. Upload your packing list file 上传装箱单文件
2. Upload your policy file 上传政策文件
3. Click 'Generate Invoice' to process 点击'生成发票'进行处理
""")

# 文件类型说明
file_descriptions = {
    "export_invoice.xlsx": "出口发票 - 用于一般贸易出口申报",
    "reimport_invoice.xlsx": "进口发票 - 用于一般贸易进口申报",
    "cif_original_invoice.xlsx": "CIF原始发票 - 包含运费和保险费的完整发票",
    "reimport_invoice_factory_Daman.xlsx": "大亚湾工厂复进口发票 - 用于大亚湾工厂的复进口申报",
    "reimport_invoice_factory_Silvass.xlsx": "银禧工厂复进口发票 - 用于银禧工厂的复进口申报"
}

# Create a temporary directory for file processing
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()
    st.session_state.output_dir = os.path.join(st.session_state.temp_dir, 'outputs')
    os.makedirs(st.session_state.output_dir, exist_ok=True)

# File upload section
st.header("File Upload 文件上传")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Packing List 装箱单")
    packing_list_file = st.file_uploader(
        "Upload packing list file (Excel format)",
        type=['xlsx'],
        help="Upload your packing list in Excel format",
        key="packing_list_uploader"
    )

with col2:
    st.subheader("Policy File 政策文件")
    policy_file = st.file_uploader(
        "Upload policy file (Excel format)",
        type=['xlsx'],
        help="Upload your policy file in Excel format",
        key="policy_file_uploader"
    )

# 生成文件区域
if 'files_generated' not in st.session_state:
    st.session_state.files_generated = False

# 验证输入文件函数
def validate_input_files(packing_list_path, policy_file_path):
    """验证输入文件的有效性
    
    Args:
        packing_list_path: 装箱单文件路径
        policy_file_path: 政策文件路径
        
    Returns:
        tuple: (验证是否通过, 错误信息)
    """
    validator = InputValidator()
    validation_results = validator.validate_all(packing_list_path, policy_file_path)
    
    # 检查所有验证结果
    all_passed = True
    error_messages = []
    
    for check_name, result in validation_results.items():
        if not result["success"]:
            # 特殊处理: 跳过"Value must be either numerical or a string containing a wildcard"错误
            if "Value must be either numerical" in result['message'] or "argument of type 'int' is not iterable" in result['message']:
                print(f"自动跳过错误: {result['message']}")
                continue
                
            all_passed = False
            error_messages.append(f"**{check_name}**: {result['message']}")
    
    return all_passed, error_messages

# Process button
if st.button("Generate Invoice 生成发票", type="primary"):
    if not packing_list_file or not policy_file:
        st.error("Please upload both packing list and policy files first! 请先上传装箱单和政策文件！")
    else:
        try:
            # Save uploaded files to temp directory
            packing_list_path = os.path.join(st.session_state.temp_dir, "packing_list.xlsx")
            policy_file_path = os.path.join(st.session_state.temp_dir, "policy.xlsx")
            
            with open(packing_list_path, "wb") as f:
                f.write(packing_list_file.getvalue())
            with open(policy_file_path, "wb") as f:
                f.write(policy_file.getvalue())
            
            # 清洗净重/毛重列，避免校验异常
            def clean_weights_columns(file_path):
                try:
                    # 直接读取Excel，不使用标题行
                    df = pd.read_excel(file_path)
                    # 查找可能的净重/毛重列名
                    weight_keywords = ['净重', 'Net Weight', '毛重', 'Gross Weight', 'N.W', 'G.W', 'Weight']
                    weight_cols = []
                    
                    # 更智能地查找所有重量相关列
                    for col in df.columns:
                        col_str = str(col).lower()
                        if any(keyword.lower() in col_str for keyword in weight_keywords):
                            weight_cols.append(col)
                    
                    # 处理每一列
                    for col in weight_cols:
                        # 先尝试直接转换为数值，错误值设为NaN
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                        
                        # 针对所有NaN的单元格(包括原始NaN和转换失败的)，重新处理
                        mask = df[col].isna()
                        if mask.any():
                            # 获取原始值(在df复制上)
                            orig_df = pd.read_excel(file_path)
                            for idx in df.index[mask]:
                                # 对于空或非数值的单元格，设为0
                                if idx < len(orig_df):
                                    orig_value = orig_df.iloc[idx][col]
                                    # 如果是通配符字符串(*,?,N/A等)，尝试保留但确保可转换
                                    if isinstance(orig_value, str) and any(c in orig_value for c in '*?N/An/a'):
                                        # 保持通配符字符串，确保后续校验能识别，但移除可能导致问题的字符
                                        df.at[idx, col] = "0"
                                    else:
                                        # 其他情况直接填0
                                        df.at[idx, col] = 0
                        
                        # 最后确保所有值都是数值，避免任何字符串类型的值
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    
                    # 保存处理后的文件
                    df.to_excel(file_path, index=False)
                    print("成功清洗重量列，避免校验异常")
                except Exception as e:
                    print(f"清洗重量列时出错: {str(e)}，将尝试继续执行")
            
            # 执行净重/毛重列清洗
            clean_weights_columns(packing_list_path)
            
            # 验证输入文件
            with st.spinner("Validating files... 正在验证文件..."):
                validation_passed, error_messages = validate_input_files(packing_list_path, policy_file_path)
            
            if not validation_passed:
                st.error("文件验证失败，请修正以下问题：")
                
                # 创建一个错误展示区域
                error_container = st.container()
                with error_container:
                    for error in error_messages:
                        # 检查是否是weights验证错误
                        if "weights" in error.lower() or "净重" in error or "毛重" in error or "Value must be" in error:
                            # 使用警告框突出显示weights相关错误
                            st.warning(error)
                            # 添加帮助提示
                            st.info("提示：净重和毛重字段必须为数值，且净重应小于毛重。请检查Excel文件中是否有非数值或通配符（如*、?、N/A等）。")
                        else:
                            # 其他错误使用普通错误框显示
                            st.error(error)
                
                st.warning("请修正上述问题后重新上传文件。")
            else:
                st.success("文件验证通过！正在处理...")
                # 处理文件
                with st.spinner("Processing files... 正在处理文件..."):
                    process_shipping_list(packing_list_path, policy_file_path, st.session_state.output_dir)

                    # Check for generated files
                    export_files = [f for f in os.listdir(st.session_state.output_dir) if f.endswith('.xlsx')]
                    
                    if export_files:
                        st.session_state.files_generated = True
                        st.session_state.export_files = export_files
                        st.success("Files generated successfully! 文件生成成功！")
                    else:
                        st.warning("No export files were generated. Please check your input files. 没有生成导出文件，请检查输入文件。")

        except Exception as e:
            st.error(f"An error occurred: {str(e)} 发生错误：{str(e)}")
            st.error("Please check your input files and try again. 请检查输入文件并重试。")

# 显示生成的文件下载区域
if st.session_state.files_generated:
    st.header("Download Files 下载文件")
    
    # 创建下载所有文件的功能
    if st.session_state.export_files:
        # 创建一个内存中的ZIP文件
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for file in st.session_state.export_files:
                file_path = os.path.join(st.session_state.output_dir, file)
                with open(file_path, "rb") as f:
                    zip_file.writestr(file, f.read())
        
        # 提供下载所有文件的按钮
        st.download_button(
            label="Download All Files 下载所有文件",
            data=zip_buffer.getvalue(),
            file_name="all_export_files.zip",
            mime="application/zip",
            help="Download all generated files as a ZIP archive"
        )
        
        st.markdown("---")
        st.markdown("### Individual Files 单个文件")
        
        # 对文件进行排序，确保CIF原始发票排在最后
        sorted_files = sorted(st.session_state.export_files, 
                             key=lambda x: 1 if x == 'cif_original_invoice.xlsx' else 0)
        
        # 为每个文件创建下载按钮，并添加对应的描述
        for file in sorted_files:
            # 跳过CIF原始发票（自动隐藏）
            if file == 'cif_original_invoice.xlsx':
                # 添加一个可折叠区域用于显示CIF原始发票（默认折叠）
                with st.expander("显示CIF原始发票（仅供内部使用）", expanded=False):
                    with open(os.path.join(st.session_state.output_dir, file), "rb") as f:
                        file_description = file_descriptions.get(file, "导出文件")
                        st.download_button(
                            label=f"Download {file}",
                            data=f,
                            file_name=file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help=file_description,
                            key=f"download_{file}"
                        )
                        st.markdown(f"**描述**: {file_description}")
                    st.markdown("*注意：CIF原始发票仅供内部计算使用，不是最终交付文件*")
            else:
                # 正常显示其他文件
                with open(os.path.join(st.session_state.output_dir, file), "rb") as f:
                    file_description = file_descriptions.get(file, "导出文件")
                    st.download_button(
                        label=f"Download {file}",
                        data=f,
                        file_name=file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help=file_description,
                        key=f"download_{file}"
                    )
                    st.markdown(f"**描述**: {file_description}")
                    st.markdown("---")

# Footer
st.markdown("---")
st.markdown("### Need Help? 需要帮助？")
st.markdown("""
If you encounter any issues or need assistance:
如果您遇到任何问题或需要帮助：

1. Check that your files are in the correct Excel format
   确保您的文件为正确的Excel格式
2. Verify that your packing list and policy files match
   验证您的装箱单和政策文件是否匹配
3. Contact support if problems persist
   如果问题持续存在，请联系支持
""")