import streamlit as st

# 1. 基礎設定 (必須放在第一行)
st.set_page_config(page_title="投保名單檢查工具", page_icon="🚄")

# 2. 安全載入套件
try:
    import pandas as pd
    import io
    import msoffcrypto
    from datetime import datetime
    import zipfile
    import xlsxwriter
    import openpyxl
except ImportError as e:
    st.error("🛑 系統錯誤：缺少必要的套件。")
    st.info("請檢查 requirements.txt 是否包含：streamlit, pandas, openpyxl, msoffcrypto-tool, XlsxWriter")
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
    """ 
    改良版解密函式：
    1. 自動偵測檔案是否有加密。
    2. 若有加密 -> 用密碼解鎖。
    3. 若無加密 -> 直接讀取 (忽略密碼)。
    """
    file_stream = io.BytesIO(file_content)
    
    try:
        office_file = msoffcrypto.OfficeFile(file_stream)
        
        # 判斷檔案是否真的被加密
        if office_file.is_encrypted():
            if not password:
                return None, False, "檔案已加密，請輸入密碼。"
            
            # 嘗試解密
            try:
                office_file.load_key(password=password)
                decrypted = io.BytesIO()
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                return decrypted, True, "OK" # True 表示原本是加密的
            except Exception:
                return None, False, "密碼錯誤，無法解鎖。"
        else:
            # 檔案沒加密，直接回傳原檔
            file_stream.seek(0)
            return file_stream, False, "OK" # False 表示原本沒加密

    except Exception as e:
        # 如果 msoffcrypto 無法讀取 (例如非 Office 檔)，嘗試直接回傳
        file_stream.seek(0)
        return file_stream, False, "OK"

def process_single_file(filename, content, password):
    # 1. 解密與讀取 (使用改良版函式)
    decrypted_stream, is_encrypted, msg = get_decrypted_stream(content, password)
    
    if decrypted_stream is None:
        return None, {"filename": filename, "status": "Fail", "msg": msg}

    # 2. 讀取 Excel 內容
    try:
        # 自動尋找表頭 (讀前30列判斷)
        preview = pd.read_excel(decrypted_stream, nrows=30, header=None)
        decrypted_stream.seek(0)
        
        header_idx = 0
        found_header = False
        for idx, row in preview.iterrows():
            row_str = row.astype(str).values
            if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
                header_idx = idx
                found_header = True
                break
        
        if not found_header:
             # 如果找不到關鍵字，嘗試直接讀第一列
             header_idx = 0

        df = pd.read_excel(decrypted_stream, header=header_idx)
    except Exception as e:
        return None, {"filename": filename, "status": "Fail", "msg": f"Excel 讀取失敗 ({str(e)})"}

    # 3. 尋找關鍵欄位
    cols = df.columns.tolist()
    id_col_name = next((c for c in cols if '身分證' in str(c)), None)
    birth_col_name = next((c for c in cols if '生日' in str(c) and '民國' in str(c)), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (含加密)"

    if not id_col_name or not birth_col_name:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位 (需包含'身分證'與'生日(民國)')"}

    # 4. 準備輸出與錯誤檢查
    output = io.BytesIO()
    error_cells = [] 
    
    id_col_idx = df.columns.get_loc(id_col_name)
    birth_col_idx = df.columns.get_loc(birth_col_name)

    for index, row in df.iterrows():
        # (A) 檢查生日
        birth_val = row[birth_col_name]
        birth_dt = parse_roc_birthday(birth_val)
        
        is_birth_error = False
        if birth_dt is None:
            stats["errors"] += 1
            error_cells.append((index, birth_col_idx))
            is_birth_error = True
        else:
            age = calculate_age(birth_dt)
            if 0 <= age < 15: stats["under_15"] += 1
            elif age >= 15: stats["adult"] += 1

        # (B) 檢查身分證
        id_val = str(row[id_col_name]).strip() if pd.notna(row[id_col_name]) else ""
        if not id_val or id_val == 'nan' or len(id_val) != 10:
             # 避免重複計算錯誤數 (如果生日已經錯了，這裡就不重複+1，但座標還是要標記)
             if not is_birth_error: 
                 stats["errors"] += 1
             error_cells.append((index, id_col_idx))

    # 5. 寫入 Excel (使用 xlsxwriter)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 標記黃底
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        for r, c in error_cells:
            value = df.iat[r, c]
            if pd.isna(value): value = ""
            worksheet.write(r + 1, c, value, yellow_format)
            
        worksheet.set_column(0, len(cols)-1, 15)

        # 6. 加密設定
        # 邏輯：原本有加密 OR 使用者有填密碼 -> 輸出就加密
        final_password = password if (is_encrypted or password) else None
        if final_password:
            workbook.set_encryption(final_password)

    output.seek(0)
    return output, stats

# ================= 網頁介面 (UI) =================
st.title("🚄 科普列車 - 投保名單自動檢查工具")
st.markdown(f"**檢查標準日：{REF_DATE.date()}**")
st.info("說明：若檔案有加密，請在左側輸入密碼。輸出之 ZIP 檔無密碼，但解壓縮後的 Excel 會自動加上密碼保護。")

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
                # 確保讀取指標歸零
                content = file.read()
                file.seek(0) 
                
                processed_data, stats = process_single_file(file.name, content, password)
                
                summary_report.append(stats)
                if processed_data:
                    processed_files.append((f"已檢查_{file.name}", processed_data))
            except Exception as e:
                st.error(f"檔案 {file.name} 發生錯誤: {str(e)}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        st.success("檢查完成！")
        st.dataframe(pd.DataFrame(summary_report))

        if processed_files:
            zip_buffer = io.BytesIO()
            # 製作標準 ZIP (Windows 可開)
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fname, f_data in processed_files:
                    zf.writestr(fname, f_data.getvalue())
                
                # 產生報告
                report_str = f"【檢查報告 {datetime.now().strftime('%H:%M')}】\n"
                for item in summary_report:
                    report_str += f"{item['filename']}: {item['msg']}\n"
                    if item['status'] == 'Success':
                         report_str += f"   - 未滿15歲: {item['under_15']}\n   - 成人: {item['adult']}\n   - 錯誤數: {item['errors']}\n"
                    report_str += "-"*20 + "\n"
                zf.writestr("總表統計.txt", report_str)

            st.download_button(
                label="📦 下載檢查結果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="檢查結果.zip",
                mime="application/zip"
            )
        else:
            st.warning("沒有成功產出的檔案，請檢查密碼或檔案內容。")
