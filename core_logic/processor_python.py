import pandas as pd
import os

def process_excel_pure_python(source_path, template_path, output_dir):
    """
    使用 Pure Python (Pandas) 处理 Excel，无需安装 Excel 软件。
    逻辑：
    1. 读取源文件 (获取链接)
    2. 读取模板文件 (获取固定文案)
    3. 循环将链接填入模板 (内存中操作)
    4. 模拟公式计算: =B2&CHAR(10)&C2&D2&" "&CHAR(10)&E2
    5. 导出结果
    """
    print(f"开始处理 (Cloud Mode)...\n源文件: {source_path}\n模板: {template_path}")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # 1. 读取源文件
        source_df = pd.read_excel(source_path)
        # 假设链接在第一列，或者找列名包含 'link'/'short'
        link_col = next((c for c in source_df.columns if 'link' in str(c).lower() or '短链' in str(c)), source_df.columns[0])
        links = source_df[link_col].dropna().tolist()
        print(f"找到 {len(links)} 个链接")

        # 2. 读取模板
        # header=0 假设第一行是表头
        template_df = pd.read_excel(template_path, header=0)
        
        # 3. 确定列索引 (根据用户公式 =B2&CHAR(10)&C2&D2&...&E2)
        # Excel A=0, B=1, C=2, D=3, E=4
        # 我们需要确保 template_df 至少有这么多列
        # 同时要保留所有列以便导出 map 到 H-L
        
        # 寻找关键列用于过滤和分组
        # "语言标识", "区域列表", "文案"
        lang_col_name = next((c for c in template_df.columns if '语言' in str(c) or 'Language' in str(c)), None)
        region_col_name = next((c for c in template_df.columns if '区域' in str(c) or 'Region' in str(c)), None)
        text_id_col_name = next((c for c in template_df.columns if '文案' in str(c) or 'Text' in str(c)), template_df.columns[0])
        
        print(f"关键列: 语言={lang_col_name}, 区域={region_col_name}, 文案ID={text_id_col_name}")
        
        master_results = []

        # 4. 循环处理每个链接
        for link in links:
            # 复制一份模板
            current_df = template_df.copy()
            
            # 填入链接到 D 列 (Index 3)
            # 注意：如果 template_df 原本没有 D 列，需要创建
            # 最好通过列名操作，但这里我们根据公式位置操作
            
            # 确保列存在，如果不够则补齐
            while len(current_df.columns) < 5:
                current_df[f'Unnamed_{len(current_df.columns)}'] = ""

            # 获取列名以便引用
            col_b = current_df.columns[1] # 正文
            col_c = current_df.columns[2] # 链接前缀?
            col_d = current_df.columns[3] # 链接坑位 (Target)
            col_e = current_df.columns[4] # 链接后缀?
            
            # 填入链接
            current_df[col_d] = link
            
            # 模拟公式: =B & \n & C & D & " " & \n & E
            # 注意处理 NaN 为空字符串
            part_b = current_df[col_b].fillna("").astype(str)
            part_c = current_df[col_c].fillna("").astype(str)
            part_d = current_df[col_d].fillna("").astype(str)
            part_e = current_df[col_e].fillna("").astype(str)
            
            # 构造内容
            # CHAR(10) is newline \n
            # Logic: B + \n + C + D + " " + \n + E
            # 注意：Excel 的 & 是强制连接
            
            # 我们可以创建一个新列 'GeneratedContent'
            generated_content = (
                part_b + "\n" + 
                part_c + 
                part_d + " " + "\n" + 
                part_e
            )
            
            # 将计算结果填入？用户说导出 H-L 列
            # 假设结果应该在原来的某个位置，或者我们需要把它放到 H-L 中的某一列
            # 由于不知道 H-L 具体是啥，我们假设 H-L 是公式计算结果所在的区域
            # H=7, I=8, J=9, K=10, L=11
            
            # 如果模板里 H 列本来就是空的或者放公式的，我们把 result 放进去？
            # 暂时我们将生成的 Content 覆盖到 H 列 (假设)，Length 覆盖到 I 列 (假设)
            # 或者我们简单地把 GeneratedContent 添加到 DataFrame，然后在导出时放在正确位置
            
            current_df['__Calculated_Content__'] = generated_content
            current_df['__Calculated_Length__'] = generated_content.str.len()
            
            # *重要猜测*: 用户的模板里，公式可能就在 H 列或 L 列？
            # 既然是 "导出 test excel 中的 H-L 列"，说明 H-L 列包含了我们需要的所有信息
            # 且 H-L 列里的数据是基于公式生成的
            # 在 Pandas 里，我们无法"覆盖"公式列让它自己变，我们必须知道公式列是哪一列
            
            # 策略：如果无法确定 H-L 具体是哪几列，我们尽量保留原数据，并将计算结果附加上
            # 但用户明确要求导出 H-L
            
            # 让我们尝试把计算结果写入到 H 列 (Index 7) 和 I 列 (Index 8)
            # 如果 DataFrame 不够宽，就补齐
            while len(current_df.columns) <= 11:
                current_df[f'Col_{len(current_df.columns)}'] = None
                
            cols = current_df.columns
            # 假设 H 列 (Index 7) 是内容? 
            # 假设 I 列 (Index 8) 是长度?
            # 这是一个风险点，但根据 "生成相应单元格内容" 和 "导出 H-L"，我们必须把计算值放进去
            
            # 我们可以把计算出的 Content 填入某个列
            # 也许 H-L 列引用了 A-E 列？
            # 如果我们只是修改了 A-E (主要是 D)，H-L 里的公式在 Python 里不会自动更新
            # 所以我们需要手动更新 H-L
            
            # 鉴于信息有限，我将把计算出的 Content 和 Length 放在最后，
            # 并同时导出 H-L (原本的列) + 新计算的列，供用户检查
            # 另外，如果 columns[7] (H) 是空的或看起来像公式，我们尝试填入
            
            # 这里做一个大胆假设：用户需要的是我们生成的内容
            # 我会把 __Calculated_Content__ 作为一个明确的列
            
            master_results.append(current_df)

        # 合并所有结果
        full_df = pd.concat(master_results, ignore_index=True)
        
        # 5. 过滤和导出
        # "当语言标识、区域列表这两列中的单元格是完整的时候"
        if lang_col_name and region_col_name:
            full_df = full_df.dropna(subset=[lang_col_name, region_col_name])
            
        # "分别导出... 文案列标注为1的... 2的"
        # 导出 H-L 列 (Index 7 to 12, exclusive) -> 7,8,9,10,11
        # 还要包含我们计算出的新列，以防 H-L 没更新
        
        # 定义导出器
        def export_group(group_id):
            subset = full_df[full_df[text_id_col_name] == group_id]
            if subset.empty:
                return

            # 导出 H-L 列
            # 必须确保我们把计算出的内容放进去替换原来的公式
            # 假设 H 列是 Content (Index 7) - 这是一个猜测
            # 为了安全，我把 H-L 列取出来，并且把 Calculated Content 附加在后面
            # 这样用户如果发现 H 列没变（因为是死公式字符串），可以用后面计算好的列
            
            target_cols = list(subset.columns[7:12]) 
            final_data = subset[target_cols].copy()
            
            # 添加计算列，方便用户
            final_data['Calculated_Content (Python)'] = subset['__Calculated_Content__']
            final_data['Calculated_Length (Python)'] = subset['__Calculated_Length__']
            
            save_path = os.path.join(output_dir, f"cloud_output_group_{group_id}.xlsx")
            final_data.to_excel(save_path, index=False)
            print(f"导出: {save_path}")

        export_group(1)
        export_group(2)
        
        print("处理完成 (Python Mode)!")

    except Exception as e:
        print(f"Error: {e}")
        raise e

if __name__ == "__main__":
    # Test
    source = r"d:/短信/20260130_海灯节/short-link-admin_download_task1391718_result.xlsx"
    template = r"d:/短信/20260130_海灯节/test.xlsx"
    out = r"d:/Antigravity/projects/output_python"
    process_excel_pure_python(source, template, out)
