import streamlit as st

# 1. 先設定頁面 (這行一定要放在最上面，不然會報錯)
st.set_page_config(page_title="投保名單檢查工具", page_icon="🚄")

# 2. 安全載入套件 (防當機檢查)
try:
    import pandas as pd
    import io
    import msoffcrypto
    from datetime import datetime
    import zipfile
    import xlsxwriter
    import openpyxl
except ImportError as e:
    st.error("🛑 網站啟動失敗！因為缺少必要的套件。")
    st.warning(f"錯誤訊息: {e}")
    st.info("請檢查您的 requirements.txt 檔案，確認裡面有包含以下內容：\n\nstreamlit\npandas\nopenpyxl\nmsoffcrypto-tool\nXlsxWriter")
    st.stop()

# ================= 設定區 =================
REF_DATE = datetime(2025, 10, 20)

# ================= 函式區 =================
def parse_roc_birthday(roc_val):
    """ 解析民國年，回傳 datetime """
    if pd.isna(roc_val): return None
    s = str(roc_val).strip().replace('\t', '').replace(' ', '')
    if s == '' or s.lower() == 'nan': return None
    s_clean = s.replace('年', '.').replace('月', '.').replace('日', '').replace('-', '.').replace('/', '.')
    
    parts = []
    if '.' in s_clean:
        parts = s_clean.split('.')
    elif s_clean.isdigit():
        if len(s_clean) == 6: parts = [s_clean[:2], s_clean[2:4], s_clean[4:]]
        elif len(s_clean) == 7: parts = [s_clean[:3], s_clean[3:5], s_clean[5:]]
    try:
        if len(parts) != 3: return None
        y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        if not (1 <= m <= 12 and 1 <= d <= 31): return None
        return datetime(y + 1911, m, d)
    except:
        return None

def calculate_age(born):
    if born is None: return -1
    return REF_DATE.year - born.year - ((REF_DATE.month, REF_DATE.day) < (born.month, born.day))

def get_decrypted_stream(file_content, password):
    """ 解密檔案，回傳 (BytesIO, 是否原本有加密) """
    file_stream = io.BytesIO(file_content)
    try:
        # 嘗試直接讀取 (如果沒加密)
        pd.read_excel(file_stream, nrows=1)
        file_stream.seek(0)
        return file_stream, False
    except:
        file_stream.seek(0)
    
    # 嘗試用密碼解鎖
    if password:
        try:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return decrypted, True
        except:
            return None, False
    return None, False

def process_single_file(filename, content, password):
    # 1. 讀取與解密
    decrypted_stream, is_encrypted = get_decrypted_stream(content, password)
    
    if decrypted_stream is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    # 2. 讀取 Excel
    try:
        # 找表頭
        preview = pd.read_excel(decrypted_stream, nrows=30, header=None)
        decrypted_stream.seek(0)
        
        header_idx = 0
        for idx, row in preview.iterrows():
            row_str = row.astype(str).values
            if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
                header_idx = idx
                break
        
        df = pd.read_excel(decrypted_stream, header=header_idx)
    except Exception as e:
        return None, {"filename": filename, "status": "Fail", "msg": f"讀取失敗: {str(e)}"}

    # 3. 找欄位
    cols = df.columns.tolist()
    id_col_name = next((c for c in cols if '身分證' in str(c)), None)
    birth_col_name = next((c for c in cols if '生日' in str(c) and '民國' in str(c)), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (已重新加密)"

    if not id_col_name or not birth_col_name:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位"}

    # 4. 準備寫入
    output = io.BytesIO()
    error_cells = [] 
    
    # 欄位索引
    id_col_idx = df.columns.get_loc(id_col_name)
    birth_col_idx = df.columns.get_loc(birth_col_name)

    # 檢查邏輯
    for index, row in df.iterrows():
        # 生日檢查
        birth_val = row[birth_col_name]
        birth_dt = parse_roc_birthday(birth_val)
        
        if birth_dt is None:
            stats["errors"] += 1
            error_cells.append((index, birth_col_idx))
        else:
            age = calculate_age(birth_dt)
            if 0 <= age < 15: stats["under_15"] += 1
            elif age >= 15: stats["adult"] += 1

        # 身分證檢查
        id_val = str(row[id_col_name]).strip() if pd.notna(row[id_col_name]) else ""
        if not id_val or id_val == 'nan' or len(id_val) != 10:
             if birth_dt is not None: 
                 stats["errors"] += 1
             error_cells.append((index, id_col_idx))

    # 5. 寫入加密 Excel (使用 XlsxWriter)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 黃底格式
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        
        # 標記黃底
        for r, c in error_cells:
            value = df.iat[r, c]
            if pd.isna(value): value = ""
            worksheet.write(r + 1, c, value, yellow_format) # +1 避開表頭
            
        worksheet.set_column(0, len(cols)-1, 15)

        # 加密設定 (關鍵)
        final_password = password if (is_encrypted or password) else None
        if final_password:
            workbook.set_encryption(final_password)

    output.seek(0)
    return output, stats

# ================= 網頁介面 (UI) =================
st.title("🚄 科普列車 - 投保名單自動檢查工具")
st.markdown(f"**檢查標準日：{REF_DATE.date()}**")
st.info("說明：若 Excel 原本有加密，處理後會自動用「原密碼」重新加密保護。下載的是 ZIP 檔，解壓縮後的 Excel 才需要密碼。")

# 側邊欄
with st.sidebar:
    st.header("⚙️ 設定")
    password = st.text_input("檔案密碼", type="password")
    st.caption("請輸入 Excel 開啟密碼 (若無則留空)。")

# 上傳區
uploaded_files = st.file_uploader("請選擇 Excel 檔案", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始檢查", type="primary"):
        progress_bar = st.progress(0)
        processed_files = []
        summary_report = []
        
        for i, file in enumerate(uploaded_files):
            try:
                content = file.read()
                file.seek(0)
                processed_data, stats = process_single_file(file.name, content, password)
                
                summary_report.append(stats)
                if processed_data:
                    processed_files.append((f"已檢查_{file.name}", processed_data))
            except Exception as e:
                st.error(f"檔案 {file.name} 發生未知錯誤: {e}")

            progress_bar.progress((i + 1) / len(uploaded_files))

        st.success("檢查完成！")
        st.dataframe(pd.DataFrame(summary_report))

        if processed_files:
            zip_buffer = io.BytesIO()
            # 製作標準 ZIP (不加密，確保 Windows 可開)
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fname, f_data in processed_files:
                    zf.writestr(fname, f_data.getvalue())
                
                # 報告
                report_str = f"【檢查報告 {datetime.now().strftime('%H:%M')}】\n"
                for item in summary_report:
                    report_str += f"{item['filename']}: {item['msg']}\n"
                zf.writestr("報告.txt", report_str)

            st.download_button(
                label="📦 下載檢查結果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="檢查結果.zip",
                mime="application/zip"
            )
        else:
            st.error("處理失敗，請檢查密碼是否正確。")
