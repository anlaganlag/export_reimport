import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
import io
from process_shipping_list import process_shipping_list, read_policy_file
from pathlib import Path

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

# Process button
if st.button("Generate Invoice 生成发票", type="primary"):
    if not packing_list_file or not policy_file:
        st.error("Please upload both packing list and policy files first! 请先上传装箱单和政策文件！")
    else:
        try:
            with st.spinner("Processing files... 正在处理文件..."):
                # Save uploaded files to temp directory
                packing_list_path = os.path.join(st.session_state.temp_dir, "packing_list.xlsx")
                policy_file_path = os.path.join(st.session_state.temp_dir, "policy.xlsx")
                
                with open(packing_list_path, "wb") as f:
                    f.write(packing_list_file.getvalue())
                with open(policy_file_path, "wb") as f:
                    f.write(policy_file.getvalue())

                # Process the files
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