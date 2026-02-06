import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# ================= 0. 系統環境檢查 =================
try:
    import openpyxl
    import msoffcrypto
    import xlsxwriter
except ImportError:
    st.error("🛑 缺少套件！請檢查 requirements.txt 是否包含: streamlit, pandas, openpyxl, msoffcrypto-tool, XlsxWriter")
    st.stop()

# ================= 1. 共用函式區 =================
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

def load_excel_robust(file_content, password):
    """ 強韌的讀取：先試密碼，再試直接開 """
    # 策略 A: 有密碼先解密
    if password:
        try:
            file_stream = io.BytesIO(file_content)
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.read_excel(decrypted, header=None), "加密解鎖成功"
        except:
            pass # 失敗就繼續往下

    # 策略 B: 直接讀取
    try:
        file_stream = io.BytesIO(file_content)
        return pd.read_excel(file_stream, header=None), "直接讀取成功"
    except:
        return None, "讀取失敗 (密碼錯誤或格式不支援)"

# ================= 2. 功能一：檢查邏輯 =================
def run_checker(uploaded_files, password):
    progress_bar = st.progress(0)
    processed_files = []
    summary_report = []
    
    for i, file in enumerate(uploaded_files):
        # 1. 讀取
        raw_df, msg = load_excel_robust(file.read(), password)
        file.seek(0)
        
        if raw_df is None:
            st.error(f"❌ {file.name}: {msg}")
            summary_report.append({"filename": file.name, "msg": msg, "status": "Fail"})
            continue

        # 2. 找表頭與整理 DataFrame
        header_idx = 0
        found_header = False
        # 讀取前 30 列找關鍵字
        for idx, row in raw_df.head(30).iterrows():
            row_str = row.astype(str).values
            if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
                header_idx = idx
                found_header = True
                break
        
        # 重整 Header
        df = raw_df.iloc[header_idx+1:].reset_index(drop=True)
        df.columns = raw_df.iloc[header_idx].values
        
        # 3. 找欄位
        cols = [str(c) for c in df.columns]
        id_col = next((c for c in cols if '身分證' in c), None)
        birth_col = next((c for c in cols if '生日' in c and '民國' in c), None)
        
        if not id_col or not birth_col:
            st.error(f"❌ {file.name}: 找不到關鍵欄位")
            summary_report.append({"filename": file.name, "msg": "找不到欄位", "status": "Fail"})
            continue

        # 4. 檢查數據
        stats = {"filename": file.name, "under_15": 0, "adult": 0, "errors": 0, "msg": "OK", "status": "Success"}
        error_cells = [] 
        
        # 取得 index
        try:
            id_idx = list(df.columns).index(id_col)
            birth_idx = list(df.columns).index(birth_col)
        except:
             st.error(f"❌ {file.name}: 欄位索引錯誤")
             continue

        for index, row in df.iterrows():
            # 生日
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

            # 身分證
            id_val = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
            if not id_val or id_val == 'nan' or len(id_val) != 10:
                if not is_birth_err: stats["errors"] += 1
                error_cells.append((index, id_idx))

        # 5. 輸出 (僅檢查，不加密輸出，確保穩定)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            yellow = workbook.add_format({'bg_color': '#FFFF00'})
            
            for r, c in error_cells:
                val = df.iat[r, c]
                if pd.isna(val): val = ""
                worksheet.write(r + 1, c, val, yellow)
        
        processed_files.append((f"已檢查_{file.name}", output.getvalue()))
        summary_report.append(stats)
        progress_bar.progress((i + 1) / len(uploaded_files))

    return processed_files, summary_report

# ================= 3. 功能二：加密邏輯 =================
def run_encryptor(uploaded_files, password):
    progress_bar = st.progress(0)
    processed_files = []
    
    for i, file in enumerate(uploaded_files):
        try:
            # 讀取 (不管原本有沒有鎖，都試著打開)
            df, msg = load_excel_robust(file.read(), None) # 這邊可以不用舊密碼，假設使用者上傳的是已檢查過(無鎖)的檔案
            # 如果上傳的是有鎖的，且沒給舊密碼，可能會失敗。
            # 但通常流程是：檢查(無鎖) -> 加密。
            
            if df is None:
                st.error(f"❌ {file.name}: 無法讀取，請確認檔案未加密或格式正確。")
                continue
                
            # 加密寫入
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # header=False 因為 load_excel_robust 是讀無 header，這裡直接寫出即可
                # 但為了美觀，建議簡單處理：
                df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                workbook = writer.book
                workbook.set_encryption(password)
            
            processed_files.append((f"加密_{file.name}", output.getvalue()))
        
        except Exception as e:
            st.error(f"❌ {file.name} 加密失敗: {e}")
            
        progress_bar.progress((i + 1) / len(uploaded_files))
        
    return processed_files

# ================= 4. 主介面 (Tabs) =================
st.set_page_config(page_title="投保名單工具箱", page_icon="🧰")
st.title("🧰 科普列車 - 投保名單工具箱")

tab1, tab2 = st.tabs(["🔍 1. 檢查名單", "🔒 2. 批次加密"])

# --- 分頁 1: 檢查 ---
with tab1:
    st.header("名單自動檢查 (年齡/身分證/黃底)")
    st.info("若檔案有加密，請輸入密碼。輸出的檔案**不會加密** (方便您確認)，確認後請至分頁 2 進行加密。")
    
    check_pass = st.text_input("輸入解鎖密碼 (若檔案無加密可留空)", type="password", key="check_pass")
    check_files = st.file_uploader("上傳 Excel 進行檢查", type=['xlsx'], accept_multiple_files=True, key="check_uploader")
    
    if check_files and st.button("🚀 開始檢查", key="btn_check"):
        results, report = run_checker(check_files, check_pass)
        
        if report:
            st.dataframe(pd.DataFrame(report))
            
        if results:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fname, data in results:
                    zf.writestr(fname, data)
                
                # 報告
                report_str = f"檢查報告 {datetime.now().strftime('%H:%M')}\n"
                for item in report:
                    if item['status'] == 'Success':
                        report_str += f"{item['filename']}: 未滿15歲: {item['under_15']}, 成人: {item['adult']}, 錯誤: {item['errors']}\n"
                    else:
                        report_str += f"{item['filename']}: {item['msg']}\n"
                zf.writestr("檢查報告.txt", report_str)
                
            st.download_button("📦 下載檢查結果 (ZIP)", zip_buffer.getvalue(), "檢查結果.zip", "application/zip")

# --- 分頁 2: 加密 ---
with tab2:
    st.header("Excel 批次加密")
    st.info("將一般的 Excel 檔案加上密碼保護。")
    
    enc_pass = st.text_input("設定新密碼 (必填)", type="password", key="enc_pass")
    enc_files = st.file_uploader("上傳要加密的 Excel", type=['xlsx'], accept_multiple_files=True, key="enc_uploader")
    
    if enc_files and enc_pass:
        if st.button("🔒 開始加密", key="btn_enc"):
            encrypted_results = run_encryptor(enc_files, enc_pass)
            
            if encrypted_results:
                st.success(f"成功加密 {len(encrypted_results)} 個檔案！")
                zip_buffer_enc = io.BytesIO()
                with zipfile.ZipFile(zip_buffer_enc, "w") as zf:
                    for fname, data in encrypted_results:
                        zf.writestr(fname, data)
                
                st.download_button("📦 下載加密檔案 (ZIP)", zip_buffer_enc.getvalue(), "已加密檔案.zip", "application/zip")
    elif enc_files and not enc_pass:
        st.warning("請輸入要設定的密碼！")
