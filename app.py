import streamlit as st
import pandas as pd
import os
import tempfile
from process_shipping_list import process_shipping_list, read_policy_file
from pathlib import Path

# Set page config
st.set_page_config(
    page_title="Export Invoice Generator",
    page_icon="📦",
    layout="wide"
)

# Title and description
st.title("Export Invoice Generator 出口发票生成器")
st.markdown("""
This application helps you generate export invoices from packing lists and policy files.
此应用程序帮助您从装箱单和政策文件生成出口发票。

### Instructions 使用说明
1. Upload your packing list file 上传装箱单文件
2. Upload your policy file 上传政策文件
3. Click 'Generate Invoice' to process 点击'生成发票'进行处理
""")

# Create a temporary directory for file processing
temp_dir = tempfile.mkdtemp()
output_dir = os.path.join(temp_dir, 'outputs')
os.makedirs(output_dir, exist_ok=True)

# File upload section
st.header("File Upload 文件上传")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Packing List 装箱单")
    packing_list_file = st.file_uploader(
        "Upload packing list file (Excel format)",
        type=['xlsx'],
        help="Upload your packing list in Excel format"
    )

with col2:
    st.subheader("Policy File 政策文件")
    policy_file = st.file_uploader(
        "Upload policy file (Excel format)",
        type=['xlsx'],
        help="Upload your policy file in Excel format"
    )

# Process button
if st.button("Generate Invoice 生成发票", type="primary"):
    if not packing_list_file or not policy_file:
        st.error("Please upload both packing list and policy files first! 请先上传装箱单和政策文件！")
    else:
        try:
            with st.spinner("Processing files... 正在处理文件..."):
                # Save uploaded files to temp directory
                packing_list_path = os.path.join(temp_dir, "packing_list.xlsx")
                policy_file_path = os.path.join(temp_dir, "policy.xlsx")
                
                with open(packing_list_path, "wb") as f:
                    f.write(packing_list_file.getvalue())
                with open(policy_file_path, "wb") as f:
                    f.write(policy_file.getvalue())

                # Process the files
                process_shipping_list(packing_list_path, policy_file_path, output_dir)

                # Check for generated files
                export_files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
                
                if export_files:
                    st.success("Files generated successfully! 文件生成成功！")
                    
                    # Create download buttons for each generated file
                    st.header("Download Files 下载文件")
                    for file in export_files:
                        with open(os.path.join(output_dir, file), "rb") as f:
                            st.download_button(
                                label=f"Download {file}",
                                data=f,
                                file_name=file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                else:
                    st.warning("No export files were generated. Please check your input files. 没有生成导出文件，请检查输入文件。")

        except Exception as e:
            st.error(f"An error occurred: {str(e)} 发生错误：{str(e)}")
            st.error("Please check your input files and try again. 请检查输入文件并重试。")

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