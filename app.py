import streamlit as st
import os
import sys
import tempfile
import io

# Import processors
sys.path.append(os.path.join(os.path.dirname(__file__), 'core_logic'))
try:
    from processor_cloud import process_excel_cloud, process_excel_cloud_get_data
except ImportError:
    sys.path.append(os.getcwd())
    from core_logic.processor_cloud import process_excel_cloud, process_excel_cloud_get_data

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

# Session State Initialization
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None

# Process Button (Step 1)
if st.button("第一步：开始分析 (Analyze)", type="primary"):
    if not uploaded_source:
        st.error("请先上传源文件！")
    elif not uploaded_template:
        st.error("请先上传模板文件！")
    else:
        try:
            with st.spinner("正在云端分析数据..."):
                # Step 1: Get data map
                data_map = process_excel_cloud_get_data(uploaded_source, uploaded_template)
                st.session_state.processed_data = data_map
                st.success(f"分析完成！共找到 {len(data_map)} 组数据。")
        except Exception as e:
            st.error(f"分析失败: {e}")
            st.exception(e)

# Rename & Download (Step 2)
if st.session_state.processed_data:
    st.markdown("---")
    st.header("3. 导出设置 (Export Configuration)")
    st.info("检测到以下分组，请依照顺序确认文件名。浏览器会自动下载到您的默认下载文件夹 (通常是 Downloads)。")
    
    # Form to collect filenames
    with st.form("filename_form"):
        renamed_files = {}
        sorted_gids = sorted(st.session_state.processed_data.keys())
        
        for gid in sorted_gids:
            group_info = st.session_state.processed_data[gid]
            default_name = group_info['default_name']
            
            col1, col2 = st.columns([1, 4])
            with col1:
                st.markdown(f"**文案组 {gid}**")
                st.caption(f"({len(group_info['data'])} 行)")
            with col2:
                new_name = st.text_input(
                    f"文件名 (文案组 {gid})", 
                    value=default_name,
                    key=f"name_{gid}",
                    help="请输入您希望保存的文件名，如 result_v1.xlsx"
                )
                if not new_name.endswith(".xlsx"):
                    new_name += ".xlsx"
                renamed_files[gid] = new_name
        
        submitted = st.form_submit_button("确认并可以下载 (Confirm & Ready)")

    # Download Buttons (Step 3)
    if submitted:
        st.success("文件名已确认！请点击下方按钮下载文件。")
        st.markdown("### ⬇️ 点击下载 (Click to Download)")
        
        for gid in sorted_gids:
            fname = renamed_files[gid]
            df = st.session_state.processed_data[gid]['data']
            
            # Convert to bytes
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            
            st.download_button(
                label=f"📥 下载: {fname}",
                data=output,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    st.caption("提示：由于网页安全限制，文件会默认保存到浏览器的下载目录中，无法直接指定保存到 D 盘某文件夹，需您手动移动。")
