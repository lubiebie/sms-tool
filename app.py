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
### 上传短链文件和初始模板文件，即可以根据文案类别自动聚合并分别导出短信模板
**完全云端运行，无需安装 Excel。**
""")

# 1. Source File Upload
st.header("1. 上传源文件 (Source)")
uploaded_source = st.file_uploader("上传短链接平台导出的短链文件", type=["xlsx", "xls"], key="source")

# 2. Template File Upload
col_t1, col_t2 = st.columns([3, 1])
with col_t1:
    st.header("2. 上传模板文件 (Template)")
    uploaded_template = st.file_uploader("请上传模板文件", type=["xlsx", "xls"], key="template")
with col_t2:
    st.write("") # Spacer
    st.write("") # Spacer
    # Read local template file to bytes
    try:
        with open("自动化工具模板.xlsx", "rb") as f:
            template_bytes = f.read()
        st.download_button(
            label="📄 点击下载模板\n(查看填写说明)",
            data=template_bytes,
            file_name="自动化工具模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except FileNotFoundError:
        st.warning("默认模板文件(自动化工具模板.xlsx)未找到")

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
        
        submitted = st.form_submit_button("确认并生成下载链接 (Confirm)")
        if submitted:
            st.session_state.confirmed_filenames = renamed_files

    # Download Buttons (Step 3) - Outside form for persistence
    if st.session_state.get('confirmed_filenames'):
        st.markdown("### ⬇️ 点击下载 (Click to Download)")
        st.success("文件名已确认！您可以直接点击下方按钮依次下载。")
        
        # Display in a grid
        cols = st.columns(3) # 3 buttons per row
        
        for idx, gid in enumerate(sorted_gids):
            fname = st.session_state.confirmed_filenames[gid]
            df = st.session_state.processed_data[gid]['data']
            
            # Convert to bytes
            output = io.BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)
            
            with cols[idx % 3]:
                st.download_button(
                    label=f"📥 {fname}",
                    data=output,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help=f"下载文案组 {gid} 的结果",
                    use_container_width=True
                )
            
    st.caption("提示：由于网页安全限制，文件会默认保存到浏览器的下载目录中，无法直接指定保存到 D 盘某文件夹，需您手动移动。")
