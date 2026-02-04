import pandas as pd
import os
import io

def process_excel_cloud(source_file, template_file, output_dir=None):
    """
    Cloud-optimized Excel processor.
    Returns a dictionary of {filename: excel_bytes} for easy download in Streamlit,
    or saves to output_dir if provided.
    """
    print("Starting Cloud Processing...")
    
    # 1. Read Source (allow file path or bytes)
    source_df = pd.read_excel(source_file)
    
    # Identify Link Column in Source
    # User said: "upload a short link excel... fill links in order"
    # Look for 'link', '短链', or use first column
    link_col = next((c for c in source_df.columns if 'link' in str(c).lower() or '链' in str(c)), source_df.columns[0])
    links = source_df[link_col].dropna().tolist()
    print(f"Source: Found {len(links)} links in column '{link_col}'")

    # 2. Read Template
    template_df = pd.read_excel(template_file)
    
    # 3. Identify Template Columns
    # Need: 正文(B), 回到提瓦特(C), 链接(D), 退订(E) -> for formula
    # Need: 语言标识, 区域列表, 发信人/签名 -> for export
    # Need: 文案 -> for grouping
    
    def find_col(keywords, default_idx=None):
        if isinstance(keywords, str): keywords = [keywords]
        for col in template_df.columns:
            for k in keywords:
                if k in str(col):
                    return col
        if default_idx is not None and default_idx < len(template_df.columns):
            return template_df.columns[default_idx]
        return None

    # Mapping based on user description + guessing standard names
    col_text_id = find_col(["文案", "Text"], 0) # Grouping key
    col_body = find_col("正文", 1)    # B
    col_back = find_col(["回到", "提瓦特", "Back"], 2) # C
    col_link_target = find_col("链接", 3)  # D (This is where we fill the link)
    col_unsub = find_col("退订", 4) # E
    
    col_lang = find_col(["语言", "Language"])
    col_region = find_col(["区域", "Region"])
    col_sender = find_col(["发信人", "签名", "Sender", "Signature"])
    col_content = find_col("内容")  # The Target Column for the formula result
    
    # Debug info
    print(f"Mapped Columns:\nBody={col_body}\nBack={col_back}\nLink={col_link_target}\nUnsub={col_unsub}\nLang={col_lang}\nRegion={col_region}\nSender={col_sender}\nContent={col_content}")
    
    if not (col_lang and col_region and col_text_id):
        raise ValueError("无法在模板中找到关键列：文案、语言标识、区域列表。请检查模板表头。")

    results = {} # Store results {filename: dataframe}

    # 4. Fill and Compute
    # Since we need to fill "in order", we repeat the template logic for each link?
    # Or does the template already have N rows, and we fill them?
    # User said: "fill links in order... generate corresponding cell content"
    # Usually this means we take the template (which might have 1 row per language) 
    # and duplicate it for each link? Or the template has matching rows?
    
    # Assumption: The template defines the structure (Multi-language pack). 
    # For *each* link in the source, we might need a whole set of languages?
    # OR, the source just provides a list of links to fill predefined slots?
    
    # "文案列标注为1的... 2的" suggests the template has multiple rows with different Copy IDs.
    # It likely defines the "Message Strategy".
    # And we just have a list of links.
    
    # Case A: One link applies to ALL template rows? (e.g. same link for all languages)
    # Case B: List of links applies 1-to-1 to template rows?
    
    # Given "short-link-admin...result", it sounds like a batch of generated links.
    # Given "Language/Region is complete", it sounds like the template is the master.
    
    # Most likely: We enter ONE link into the template (which generates content for many languages), 
    # then export. Then repeat for next link?
    # NO, "fill links into another excel... order" usually implies 1-to-1 or filling a column.
    
    # Let's assume we fill the links into `col_link_target` of the template.
    # If source has 100 links and template has 100 rows, 1-to-1.
    # If source has 1 link and template has 20 rows (languages), maybe fill same link?
    
    # User said: "fill ... in order to another excel".
    # I will assume 1-to-1 filling. If template runs out of rows, stop. 
    # If links run out, stop.
    
    # Or maybe the template is just 2 rows (ID 1 and ID 2) and we need to generate a massive file?
    # "文案列标注为1的... 导入为新的一份... 2的... 导入到第二个"
    # This implies splitting the result by ID.
    
    # Let's try filling the `col_link_target` column with the `links` list.
    filled_df = template_df.copy()
    
    # Ensure dataframe is long enough?
    if len(links) > len(filled_df):
        print("Warning: Source has more links than Template has rows. Truncating source.")
        links = links[:len(filled_df)]
    elif len(links) < len(filled_df):
         print("Warning: Source has fewer links than Template. Some rows will be empty.")
    
    # Update the Link Column
    # We only update rows where we have links.
    filled_df.loc[:len(links)-1, col_link_target] = links
    
    # 5. Compute Formula: =B2&CHAR(10)&C2&D2&" "&CHAR(10)&E2
    # Vectorized computation
    def get_str(col):
        if col: return filled_df[col].fillna("").astype(str)
        return pd.Series([""] * len(filled_df))

    b_val = get_str(col_body)
    c_val = get_str(col_back)
    d_val = get_str(col_link_target) # The links we just filled
    e_val = get_str(col_unsub)
    
    # Formula logic
    # B + \n + C + D + " " + \n + E
    newline = "\n"
    computed_content = b_val + newline + c_val + d_val + " " + newline + e_val
    
    # Store result in `col_content`
    if col_content:
        filled_df[col_content] = computed_content
    else:
        filled_df["Content_Calculated"] = computed_content
        col_content = "Content_Calculated"
        
    # 6. Filter and Export
    # Condition: Language & Region must be complete (not null)
    valid_rows = filled_df.dropna(subset=[col_lang, col_region])
    
    # Group by Text ID (文案)
    groups = valid_rows[col_text_id].unique()
    
    generated_files = {} # path -> dataframe
    
    for gid in groups:
        subset = valid_rows[valid_rows[col_text_id] == gid]
        if subset.empty: continue
        
        # Columns to export: Language, Region, Sender, Content
        export_cols = [c for c in [col_lang, col_region, col_sender, col_content] if c is not None]
        
        final_data = subset[export_cols]
        
        # Determine filename
        fname = f"output_group_{gid}.xlsx"
        if output_dir:
            fpath = os.path.join(output_dir, fname)
            final_data.to_excel(fpath, index=False)
            generated_files[fname] = fpath
        else:
            # Memory mode for web download
            output = io.BytesIO()
            final_data.to_excel(output, index=False)
            output.seek(0)
            generated_files[fname] = output
            
    return generated_files

if __name__ == "__main__":
    # Test
    src = r"d:/短信/20260130_海灯节/short-link-admin_download_task1391718_result.xlsx"
    tpl = r"d:/短信/20260130_海灯节/test.xlsx"
    out = r"d:/Antigravity/projects/output_cloud"
    process_excel_cloud(src, tpl, out)
