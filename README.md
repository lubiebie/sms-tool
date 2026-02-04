# Excel Automation Tool (Web Version)

这是一个自动化的 Excel 处理工具，用于批量生成短链内容。

## 功能
- **批量填入**：将短链列表自动填入模板。
- **自动计算**：云端模拟 Excel 公式 `=正文 & \n & ... & 退订`。
- **智能导出**：根据“文案ID”自动拆分并导出所需列。

## 如何部署 (Streamlit Cloud)
1. 将本项目所有文件上传到 GitHub。
2. 访问 [Streamlit Cloud](https://share.streamlit.io/)。
3. 选择本仓库，设置 Main file path 为 `app.py`。
4. 点击 Deploy。

## 本地运行
```bash
pip install -r requirements.txt
streamlit run app.py
```
