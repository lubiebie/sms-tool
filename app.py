import streamlit as st
import os
import sys
import tempfile
import io

# Import processors
sys.path.append(os.path.join(os.path.dirname(__file__), 'core_logic'))
try:
    from processor_cloud import process_excel_cloud
except ImportError:
    sys.path.append(os.getcwd())
    from core_logic.processor_cloud import process_excel_cloud

st.set_page_config(page_title="Excel Auto-Processing Tool", layout="wide")

st.title("📊 Excel 自动化处理工具 (Cloud)")
st.markdown("""
本工具用于将短链数据填入模板，自动计算并导出结果。
**完全云端运行，无需安装 Excel。**
""")

# 1. Source File Upload
st.header("1. 上传源文件 (Source)")
uploaded_source = st.file_uploader("上传包含短链的 Excel 文件", type=["xlsx", "xls"], key="source")

# 2. Template File Upload
st.header("2. 上传模板文件 (Template)")
uploaded_template = st.file_uploader("上传模板 Excel 文件 (包含公式和文案规则)", type=["xlsx", "xls"], key="template")

# Process Button
if st.button("开始处理 (Start Processing)", type="primary"):
    if not uploaded_source:
        st.error("请先上传源文件！")
    elif not uploaded_template:
        st.error("请先上传模板文件！")
    else:
        try:
            with st.spinner("正在云端处理..."):
                # Run processing in memory
                results = process_excel_cloud(uploaded_source, uploaded_template)
            
            st.success("处理完成！请下载结果文件：")
            
            # Display Download Buttons
            for fname, data in results.items():
                if isinstance(data, io.BytesIO): 
                    # Only handling BytesIO (Memory mode)
                    st.download_button(
                        label=f"⬇️ 下载 {fname}",
                        data=data,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                     # Fallback if it returned paths (shouldn't happen with updated logic)
                     st.write(f"文件已保存: {fname}")

        except Exception as e:
            st.error(f"处理失败: {e}")
            st.exception(e)
