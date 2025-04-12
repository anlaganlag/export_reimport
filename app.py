import streamlit as st
import pandas as pd
import os
import traceback
import base64
from process_shipping_list import process_shipping_list

def main():
    st.set_page_config(page_title="Invoice Generator", layout="wide")
    
    # Initialize session state to store generated files
    if 'generated_files' not in st.session_state:
        st.session_state.generated_files = {}
    
    st.title("Export/Reimport Invoice Generator")
    
    # Create columns for better layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.header("Required Files")
        packing_list = st.file_uploader("Upload Packing List", type=['xlsx'])
        policy_file = st.file_uploader("Upload Policy File", type=['xlsx'])

    with col2:
        if packing_list and policy_file:
            st.success("✅ All required files uploaded")
        else:
            st.warning("⚠️ Please upload all required files")
    
    # Template Selection Section
    st.header("Template Selection")
    
    # Create tabs for Export and Reimport templates
    export_tab, reimport_tab = st.tabs(["Export Invoice Templates", "Reimport Invoice Templates"])
    
    with export_tab:
        st.subheader("Export Invoice Templates")
        
        # Commercial Invoice templates
        st.markdown("**Commercial Invoice:**")
        exp_ci_template = st.radio(
            "Choose Commercial Invoice Template Source:",
            ["Use Default", "Upload Custom"],
            key="exp_ci"
        )
        
        if exp_ci_template == "Upload Custom":
            exp_ci_header = st.file_uploader("Upload Custom Header", type=['xlsx'], key="exp_ci_h")
            exp_ci_footer = st.file_uploader("Upload Custom Footer", type=['xlsx'], key="exp_ci_f")
        else:
            st.info("Using default h.xlsx and f.xlsx")
            exp_ci_header = None
            exp_ci_footer = None
        
        # Packing List templates
        st.markdown("**Packing List:**")
        exp_pl_template = st.radio(
            "Choose Packing List Template Source:",
            ["Use Default", "Upload Custom"],
            key="exp_pl"
        )
        
        if exp_pl_template == "Upload Custom":
            exp_pl_header = st.file_uploader("Upload Custom Header", type=['xlsx'], key="exp_pl_h")
            exp_pl_footer = st.file_uploader("Upload Custom Footer", type=['xlsx'], key="exp_pl_f")
        else:
            st.info("Using default pl_h.xlsx and pl_f.xlsx")
            exp_pl_header = None
            exp_pl_footer = None
    
    with reimport_tab:
        st.subheader("Reimport Invoice Templates")
        
        # Commercial Invoice templates
        st.markdown("**Commercial Invoice:**")
        reimp_ci_template = st.radio(
            "Choose Commercial Invoice Template Source:",
            ["Use Default", "Upload Custom"],
            key="reimp_ci"
        )
        
        if reimp_ci_template == "Upload Custom":
            reimp_ci_header = st.file_uploader("Upload Custom Header", type=['xlsx'], key="reimp_ci_h")
            reimp_ci_footer = st.file_uploader("Upload Custom Footer", type=['xlsx'], key="reimp_ci_f")
        else:
            st.info("Using default h.xlsx and f.xlsx")
            reimp_ci_header = None
            reimp_ci_footer = None
        
        # Packing List templates
        st.markdown("**Packing List:**")
        reimp_pl_template = st.radio(
            "Choose Packing List Template Source:",
            ["Use Default", "Upload Custom"],
            key="reimp_pl"
        )
        
        if reimp_pl_template == "Upload Custom":
            reimp_pl_header = st.file_uploader("Upload Custom Header", type=['xlsx'], key="reimp_pl_h")
            reimp_pl_footer = st.file_uploader("Upload Custom Footer", type=['xlsx'], key="reimp_pl_f")
        else:
            st.info("Using default pl_h.xlsx and pl_f.xlsx")
            reimp_pl_header = None
            reimp_pl_footer = None
    
    # Generate Button
    st.markdown("---")
    if st.button("Generate Invoices", type="primary", use_container_width=True):
        if not packing_list or not policy_file:
            st.error("Please upload both Packing List and Policy files!")
            return
            
        try:
            with st.spinner("Processing files and generating invoices..."):
                # Create outputs directory if it doesn't exist
                os.makedirs("outputs", exist_ok=True)
                
                # Save uploaded files to temporary location
                temp_packing_list = save_uploaded_file(packing_list)
                temp_policy_file = save_uploaded_file(policy_file)
                
                st.info(f"Temporary files created: {temp_packing_list}, {temp_policy_file}")
                
                # Save uploaded templates to temporary files if provided
                template_paths = {
                    'export_ci_header': save_uploaded_file(exp_ci_header) if exp_ci_header else 'h.xlsx',
                    'export_ci_footer': save_uploaded_file(exp_ci_footer) if exp_ci_footer else 'f.xlsx',
                    'export_pl_header': save_uploaded_file(exp_pl_header) if exp_pl_header else 'pl_h.xlsx',
                    'export_pl_footer': save_uploaded_file(exp_pl_footer) if exp_pl_footer else 'pl_f.xlsx',
                    'reimport_ci_header': save_uploaded_file(reimp_ci_header) if reimp_ci_header else 'h.xlsx',
                    'reimport_ci_footer': save_uploaded_file(reimp_ci_footer) if reimp_ci_footer else 'f.xlsx',
                    'reimport_pl_header': save_uploaded_file(reimp_pl_header) if reimp_pl_header else 'pl_h.xlsx',
                    'reimport_pl_footer': save_uploaded_file(reimp_pl_footer) if reimp_pl_footer else 'pl_f.xlsx'
                }
                
                # Log what templates are being used
                st.info(f"Using templates: {template_paths}")
                
                # Process files and generate invoices
                st.info("Starting processing...")
                processing_successful = process_shipping_list(
                    temp_packing_list,
                    temp_policy_file,
                    template_paths=template_paths
                )
                st.info(f"Processing completed with result: {processing_successful}")
                
                # Define files to download
                files_to_download = [
                    ('export_invoice.xlsx', 'Export Invoice'),
                    ('reimport_invoice.xlsx', 'Reimport Invoice'),
                    ('cif_original_invoice.xlsx', 'CIF Original Invoice')
                ]
                
                # Check if files were created
                existing_files = [file for file, label in files_to_download if os.path.exists(f"outputs/{file}")]
                st.info(f"Files found: {existing_files}")
                
                output_files_exist = len(existing_files) > 0
                
                if output_files_exist:
                    st.success("✅ Invoices generated successfully!")
                    
                    # Display download section for generated files
                    st.subheader("Download Generated Invoices")
                    
                    # Create columns for download buttons
                    col1, col2, col3 = st.columns(3)
                    cols = [col1, col2, col3]
                    
                    # Create download buttons for each file
                    for i, (file, label) in enumerate(files_to_download):
                        if os.path.exists(f"outputs/{file}"):
                            # Add download button for this file
                            cols[i].download_button(
                                label=f"📥 Download {label}",
                                data=open(f"outputs/{file}", "rb"),
                                file_name=file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{file}",
                                use_container_width=True
                            )
                else:
                    st.warning("No output files were generated. Please check if templates and input files are correct.")
        
        except Exception as e:
            st.error(f"Error generating invoices: {str(e)}")
            st.error(traceback.format_exc())
        finally:
            # Cleanup temporary files
            cleanup_temp_files()
    
    # Always display previously generated files if they exist
    if os.path.exists("outputs"):
        files_available = False
        files_to_check = [
            ('export_invoice.xlsx', 'Export Invoice'),
            ('reimport_invoice.xlsx', 'Reimport Invoice'),
            ('cif_original_invoice.xlsx', 'CIF Original Invoice')
        ]
        
        # Check if any files exist
        for file, _ in files_to_check:
            if os.path.exists(f"outputs/{file}"):
                files_available = True
                break
        
        if files_available:
            st.markdown("---")
            st.subheader("Available Files")
            
            # Create columns for download buttons
            col1, col2, col3 = st.columns(3)
            cols = [col1, col2, col3]
            
            # Create download buttons for each available file
            for i, (file, label) in enumerate(files_to_check):
                if os.path.exists(f"outputs/{file}"):
                    # Add download button for this file
                    cols[i].download_button(
                        label=f"📥 Download {label}",
                        data=open(f"outputs/{file}", "rb"),
                        file_name=file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"available_{file}",
                        use_container_width=True
                    )

def save_uploaded_file(uploaded_file):
    """Save uploaded file to temporary location and return path."""
    if uploaded_file:
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getvalue())
        return temp_path
    return None

def cleanup_temp_files():
    """Remove any temporary files."""
    for file in os.listdir():
        if file.startswith("temp_") and file.endswith(".xlsx"):
            try:
                os.remove(file)
            except Exception as e:
                st.warning(f"Could not remove temporary file {file}: {str(e)}")

if __name__ == "__main__":
    main() 