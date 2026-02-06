import streamlit as st

# 1. 基礎設定
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
    暴力嘗試法：
    1. 有密碼 -> 先試著用密碼解。
    2. 解不開/沒密碼 -> 試著直接開。
    """
    # 策略 A: 如果使用者有給密碼，先嘗試解密
    if password:
        try:
            file_stream = io.BytesIO(file_content)
            office_file = msoffcrypto.OfficeFile(file_stream)
            
            # 準備解密
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            
            # 驗證解密後能不能讀
            decrypted.seek(0)
            pd.read_excel(decrypted, nrows=1) # 試讀一行
            decrypted.seek(0)
            
            return decrypted, True, "OK" # 成功用密碼解開
        except:
            # 密碼解鎖失敗，可能是：密碼錯、或者檔案根本沒加密
            pass # 默默失敗，進入策略 B

    # 策略 B: 嘗試直接打開 (當作沒加密)
    try:
        file_stream = io.BytesIO(file_content)
        pd.read_excel(file_stream, nrows=1) # 試讀一行
        file_stream.seek(0)
        
        # 能直接開，代表沒加密 (就算使用者有輸密碼，我們也當作 False，因為檔案本身沒鎖)
        return file_stream, False, "OK"
    except:
        pass

    # 策略 C: 全都失敗
    if password:
        return None, False, "無法讀取 (密碼錯誤，或檔案格式不支援)"
    else:
        return None, False, "無法讀取 (若是加密檔，請輸入密碼)"

def process_single_file(filename, content, password):
    # 1. 取得檔案串流
    decrypted_stream, is_encrypted, msg = get_decrypted_stream(content, password)
    
    if decrypted_stream is None:
        return None, {"filename": filename, "status": "Fail", "msg": msg}

    # 2. 讀取 Excel
    try:
        # 找表頭
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
        
        if not found_header: header_idx = 0

        df = pd.read_excel(decrypted_stream, header=header_idx)
    except Exception as e:
        return None, {"filename": filename, "status": "Fail", "msg": f"讀取失敗 ({str(e)})"}

    # 3. 欄位對應
    cols = df.columns.tolist()
    id_col_name = next((c for c in cols if '身分證' in str(c)), None)
    birth_col_name = next((c for c in cols if '生日' in str(c) and '民國' in str(c)), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (含加密)"

    if not id_col_name or not birth_col_name:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位"}

    # 4. 檢查與記錄錯誤
    output = io.BytesIO()
    error_cells = [] 
    
    id_col_idx = df.columns.get_loc(id_col_name)
    birth_col_idx = df.columns.get_loc(birth_col_name)

    for index, row in df.iterrows():
        # 生日
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

        # 身分證
        id_val = str(row[id_col_name]).strip() if pd.notna(row[id_col_name]) else ""
        if not id_val or id_val == 'nan' or len(id_val) != 10:
             if not is_birth_error: stats["errors"] += 1
             error_cells.append((index, id_col_idx))

    # 5. 寫入與加密輸出
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
        
        for r, c in error_cells:
            value = df.iat[r, c]
            if pd.isna(value): value = ""
            worksheet.write(r + 1, c, value, yellow_format)
            
        worksheet.set_column(0, len(cols)-1, 15)

        # 只要原本是加密的，或者使用者現在有填密碼，輸出就加密
        final_password = password if (is_encrypted or password) else None
        if final_password:
            workbook.set_encryption(final_password)

    output.seek(0)
    return output, stats

# ================= 網頁介面 (UI) =================
st.title("🚄 科普列車 - 投保名單自動檢查工具")
st.markdown(f"**檢查標準日：{REF_DATE.date()}**")
st.info("說明：請在左側輸入密碼。系統會自動嘗試解鎖並檢查。")

# 側邊欄
with st.sidebar:
    st.header("⚙️ 設定")
    password = st.text_input("檔案密碼", type="password")
    st.caption("請輸入 Excel 開啟密碼 (若無則留空)。")

# 上傳區
uploaded_files = st.file_uploader("請選擇 Excel 檔案", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始檢查", type="primary"):
