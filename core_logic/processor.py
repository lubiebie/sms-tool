import xlwings as xw
import pandas as pd
import os
import time

def process_excel(source_path, template_path, output_dir):
    """
    处理Excel文件的核心逻辑：
    1. 读取源文件
    2. 填入模板
    3. 等待公式计算
    4. 导出特定列
    """
    print(f"开始处理...\n源文件: {source_path}\n模板: {template_path}\n输出: {output_dir}")
    
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    app = xw.App(visible=False) #设置 visible=True 可以看到操作过程，调试时很有用
    try:
        # 打开源文件和模板文件
        # 注意：这里假设源文件第一页是数据
        source_wb = app.books.open(source_path)
        source_sheet = source_wb.sheets[0]
        
        # 读取源数据
        # 假设第一行是表头
        source_data = source_sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        
        # 根据用户描述寻找关键列
        # "文案" (Text), "语言标识" (Language ID), "区域列表" (Region List)
        # 这里需要根据实际列名进行调整，暂时使用模糊匹配或者假设
        
        # 打印列名帮助调试
        print("源文件列名:", source_data.columns.tolist())
        
        # 定义关键列名映射 (需要用户确认或自动识别)
        # 假设源文件有一列叫 "short_link" 或者类似的，需要填入模板
        # 假设模板文件需要填入的列位置
        
        wb = app.books.open(template_path)
        sheet = wb.sheets[0]
        
        processed_count = 0
        
        # 遍历源数据行
        for index, row in source_data.iterrows():
            # 1. 填入链接到模板
            # 假设源文件有一列叫 'link' 或 'short_link'，填入模板的 A 列 (假设)
            # 用户描述："将这个文件中的链接按顺序填入到另一个excel中"
            # 我们需要知道具体填入哪一列，这里暂时假设填入模板的对应位置
            
            # 假设源文件链接列名为 '短链' 或 'Short Link'，如果没有找到可以尝试第一列
            link_col = next((c for c in source_data.columns if '链' in str(c) or 'link' in str(c).lower()), source_data.columns[0])
            link_value = row[link_col]
            
            # 假设填入模板的第2行开始（第1行表头），为了简单，可以在模板的指定位置一行行填？
            # 或者用户意图是：模板是一个计算器，填入一行，计算一行，导出一行？
            # "填入之后，应该会根据test excel中设置的公式等，生成相应单元格的内容" -> 暗示是一行行处理的模式，或者是一个列表模式
            
            # 模式A：模板是一个列表，把源文件所有链接拷进去，然后公式自动拉下来
            # 模式B：模板是单行计算器，填入一个值，生成结果，保存，再填下一个
            
            # 根据 "导出test excel中的H-L列... 导入为新的一份excel"
            # 看起来更像是批量填入，然后根据结果导出
            
            # 假设模式A：批量填入
            # 找到模板中需要填入链接的列，假设是 A 列
            # 获取当前已有数据的最后一行，或者从第2行开始覆盖
            
            # 由于不确定具体位置，我将把数据填入 A 列 (从 A2 开始)，如果用户有特殊要求需修改
            current_row = index + 2 
            sheet.range(f'A{current_row}').value = link_value
            
            # 2. 等待公式计算
            # 这一步如果是批量填入，可以最后一起算
            # 如果是单行依赖（比如有些公式依赖上一行），则需要逐行
            
        # 填完所有数据后，进行一次计算（通常 Excel 会自动计算，但并未保存）
        wb.app.calculate()
        
        # 3. 检查条件并导出
        # 读取模板中计算后的数据 (包含H-L列)
        # 读取范围：假设数据在 A:L 区域
        calculated_data = sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        
        # 寻找关键列：语言标识、区域列表、文案列
        # 假设列名包含关键字
        lang_col = next((c for c in calculated_data.columns if '语言' in str(c) or 'Language' in str(c)), None)
        region_col = next((c for c in calculated_data.columns if '区域' in str(c) or 'Region' in str(c)), None)
        text_id_col = next((c for c in calculated_data.columns if '文案' in str(c) or 'Text' in str(c)), None)
        
        if not (lang_col and region_col and text_id_col):
             print(f"警告：无法在模板中找到关键列 (语言, 区域, 文案)。现有列名: {calculated_data.columns.tolist()}")
             # 尝试硬编码列索引 H-L (即第 8 到 12 列)
             # H=8, I=9, J=10, K=11, L=12
             # 假设 '文案' 是第一列 (A列)? 用户说 "根据文案（第一列）的序号"
             text_id_col = calculated_data.columns[0] 
        
        # 根据文案ID分组导出
        # 分组 1
        group1 = calculated_data[calculated_data[text_id_col] == 1]
        
        # 分组 2
        group2 = calculated_data[calculated_data[text_id_col] == 2]
        
        # 导出 H-L 列
        # H 是第 7 (0-indexed) -> L 是第 11
        # 也可以直接用列名如果知道的话
        # 用户明确说 H-L 列
        
        def save_subset(df, group_name):
            if df.empty:
                print(f"分组 {group_name} 为空，跳过。")
                return
                
            # 提取 H-L 列 (iloc 7:12)
            subset = df.iloc[:, 7:12] 
            
            # 此时还需要检查 "语言标识、区域列表" 是否完整
            # 假设这两列就在 H-L 之间，或者在之前的列？
            # "当语言标识、区域列表这两列中的单元格是完整的时候"
            # 这听起来像是一个过滤条件：只有这两列不为空的行才导出？
            
            # 再次尝试确认识别这两列
            valid_rows = df.copy()
            if lang_col and region_col:
                valid_rows = valid_rows.dropna(subset=[lang_col, region_col])
            
            final_subset = valid_rows.iloc[:, 7:12]
            
            output_path = os.path.join(output_dir, f"output_group_{group_name}.xlsx")
            final_subset.to_excel(output_path, index=False)
            print(f"已保存分组 {group_name} 到 {output_path}")

        save_subset(group1, "1")
        save_subset(group2, "2")
        
        print("处理完成！")

    except Exception as e:
        print(f"发生错误: {e}")
        raise e
    finally:
        # 关闭文件，释放资源
        # wb.close() # 调试时不关闭以便查看
        # source_wb.close()
        app.quit()

if __name__ == "__main__":
    # 简单的测试桩
    source = r"d:/短信/20260130_海灯节/short-link-admin_download_task1391718_result.xlsx"
    template = r"d:/短信/20260130_海灯节/test.xlsx"
    out = r"d:/Antigravity/projects/output"
    process_excel(source, template, out)
