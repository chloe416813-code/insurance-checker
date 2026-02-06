import streamlit as st
import io
import zipfile
from datetime import datetime

# ================= 0. 系統環境檢查 =================
# 這是為了防止網頁直接白畫面或當機
try:
    import pandas as pd
    import openpyxl
    import msoffcrypto
    import xlsxwriter
except ImportError as e:
    st.error("🛑 網頁啟動失敗！")
    st.warning(f"缺少套件: {e}")
    st.info("請確認 requirements.txt 內包含: streamlit, pandas, openpyxl, msoffcrypto-tool, XlsxWriter")
    st.stop()

# ================= 1. 核心邏輯區 =================
REF_DATE = datetime(2025, 10, 20)

def parse_roc_birthday(roc_val):
    """ 解析民國年生日 """
    if pd.isna(roc_val): return None
    s = str(roc_val).strip().replace('\t', '').replace(' ', '')
    if s == '' or s.lower() == 'nan': return None
    s_clean = s.replace('年', '.').replace('月', '.').replace('日', '').replace('-', '.').replace('/', '.')
    
    parts = []
    if '.' in s_clean: parts = s_clean.split('.')
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

def load_excel_safe(file_content, password):
    """
    超級強韌的讀取函式：
    1. 先試著直接用 openpyxl 開 (針對無加密檔案)。
    2. 失敗的話，假設是加密檔，用 msoffcrypto 解鎖。
    """
    # 嘗試 1: 直接開
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
        return wb, False, "OK" # False 代表原本沒加密
    except Exception:
        # 失敗了，可能是加密檔，進入嘗試 2
        pass

    # 嘗試 2: 用密碼解密
    if password:
        try:
            file_stream = io.BytesIO(file_content)
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            
            wb = openpyxl.load_workbook(decrypted, data_only=True)
            return wb, True, "OK" # True 代表原本是加密的
        except Exception as e:
            return None, False, "密碼錯誤或解密失敗"
    
    return None, False, "無法開啟 (檔案已加密但未輸入密碼，或檔案損毀)"

def process_single_file(filename, content, password):
    # 讀取 Excel (取得 Workbook 物件)
    wb, is_encrypted, msg = load_excel_safe(content, password)
    
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": msg}

    ws = wb.active
    
    # 將資料轉為 DataFrame 以便處理
    data = list(ws.values)
    if not data:
        return None, {"filename": filename, "status": "Fail", "msg": "檔案是空的"}

    # 尋找表頭 (讀前 30 列)
    header_idx = 0
    df = None
    
    # 簡單的表頭搜尋
    for i, row in enumerate(data[:30]):
        row_str = [str(c) if c else '' for c in row]
        if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
            header_idx = i
            break
    
    # 建立 DataFrame
    cols = data[header_idx]
    rows = data[header_idx+1:]
    df = pd.DataFrame(rows, columns=cols)

    # 尋找關鍵欄位名稱
    col_names = [str(c) for c in df.columns]
    id_col = next((c for c in col_names if '身分證' in c), None)
    birth_col = next((c for c in col_names if '生日' in c and '民國' in c), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (含加密)"

    if not id_col or not birth_col:
        return None, {"filename": filename, "status": "Fail", "msg": f"找不到欄位 (需有身分證、生日(民國))"}

    # 準備輸出
    output = io.BytesIO()
    error_cells = [] # 紀錄 (row_idx, col_idx)

    # 取得欄位索引
    id_idx = df.columns.get_loc(id_col)
    birth_idx = df.columns.get_loc(birth_col)

    for index, row in df.iterrows():
        # 1. 檢查生日
        birth_val = row[birth_col]
        birth_dt = parse_roc_birthday(birth_val)
        is_birth_err = False

        if birth_dt is None:
            stats["errors"] += 1
            error_cells.append((index, birth_idx))
            is_birth_err = True
        else:
            age = calculate_age(birth_dt)
            if 0 <= age < 15: stats["under_15"] += 1
            elif age >= 15: stats["adult"] += 1

        # 2. 檢查身分證
        id_val = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
        if not id_val or id_val == 'nan' or len(id_val) != 10:
            if not is_birth_err: stats["errors"] += 1
            error_cells.append((index, id_idx))

    # 寫入 Excel (使用 xlsxwriter)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 標記黃底
        yellow = workbook.add_format({'bg_color': '#FFFF00'})
        
        for r, c in error_cells:
            val = df.iat[r, c]
            if pd.isna(val): val = ""
            # r+1 是因為有表頭
            worksheet.write(r + 1, c, val, yellow)

        # 加密設定 (如果有密碼，就鎖回去)
        final_pass = password if (is_encrypted or password) else None
        if final_pass:
            workbook.set_encryption(final_pass)

    output.seek(0)
    return output, stats

# ================= 2. 網頁介面區 =================
st.title("🚄 科普列車 - 檢查工具 (除錯版)")
st.info("此版本會顯示詳細錯誤，請上傳檔案測試。")

# 側邊欄
with st.sidebar:
    st.header("⚙️ 設定")
    password = st.text_input("檔案密碼", type="password")

# 上傳
uploaded_files = st.file_uploader("請上傳 Excel", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始檢查", type="primary"):
        progress_bar = st.progress(0)
        processed_files = []
        summary_report = []
        
        for i, file in enumerate(uploaded_files):
            try:
                content = file.read()
                processed_data, stats = process_single_file(file.name, content, password)
                
                summary_report.append(stats)
                if processed_data:
                    processed_files.append((f"已檢查_{file.name}", processed_data))
                else:
                    # 如果失敗，顯示紅字錯誤
                    st.error(f"❌ {file.name} 失敗: {stats['msg']}")

            except Exception as e:
                st.error(f"❌ {file.name} 發生系統錯誤: {str(e)}")
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        # 顯示結果表
        if summary_report:
            st.write("### 檢查結果統計")
            st.dataframe(pd.DataFrame(summary_report))

        # 打包下載
        if processed_files:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fname, f_data in processed_files:
                    zf.writestr(fname, f_data.getvalue())
                
                report_str = "檢查報告\n" + "-"*20 + "\n"
                for item in summary_report:
                    report_str += f"{item['filename']}: {item['msg']}\n"
                zf.writestr("report.txt", report_str)

            st.download_button(
                label="📦 下載檢查結果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="檢查結果.zip",
                mime="application/zip"
            )
