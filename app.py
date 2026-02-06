import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime

# ================= 0. 系統環境防呆 =================
try:
    import openpyxl
    import msoffcrypto
    import xlsxwriter
except ImportError as e:
    st.error(f"🛑 缺少必要套件: {e}")
    st.info("請確認 requirements.txt 包含: streamlit, pandas, openpyxl, msoffcrypto-tool, XlsxWriter")
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

def get_decrypted_stream(file_content, password):
    """ 
    【經典暴力解鎖法】 
    這是之前測試最成功的版本：
    1. 有密碼 -> 優先嘗試解密。
    2. 失敗或無密碼 -> 嘗試直接開啟。
    """
    # 策略 A: 嘗試用密碼解密
    if password:
        try:
            file_stream = io.BytesIO(file_content)
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            
            # 測試是否真的解開了 (試讀一行)
            pd.read_excel(decrypted, nrows=1) 
            decrypted.seek(0)
            return decrypted, True, "OK" # True = 原本是加密的
        except:
            pass # 密碼錯誤或根本沒加密，默默失敗，換下一招

    # 策略 B: 嘗試直接開啟 (針對無加密檔案)
    try:
        file_stream = io.BytesIO(file_content)
        pd.read_excel(file_stream, nrows=1)
        file_stream.seek(0)
        return file_stream, False, "OK" # False = 原本沒加密
    except:
        pass

    # 策略 C: 都失敗
    return None, False, "無法讀取 (可能是密碼錯誤，或檔案損毀)"

# ================= 2. 分頁功能實作 =================

def run_checker(uploaded_files, password):
    """ 分頁 1: 檢查功能 (回歸最原始版本) """
    processed_files = []
    summary_report = []
    progress_bar = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        # 1. 取得檔案串流
        content = file.read()
        decrypted_stream, is_encrypted, msg = get_decrypted_stream(content, password)
        
        if decrypted_stream is None:
            # 記錄失敗
            summary_report.append({"filename": file.name, "status": "Fail", "msg": msg})
            continue

        # 2. 讀取 DataFrame
        try:
            # 找表頭
            preview = pd.read_excel(decrypted_stream, nrows=30, header=None)
            decrypted_stream.seek(0)
            
            header_idx = 0
            found = False
            for idx, row in preview.iterrows():
                row_str = row.astype(str).values
                if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
                    header_idx = idx
                    found = True
                    break
            if not found: header_idx = 0
            
            df = pd.read_excel(decrypted_stream, header=header_idx)
            
        except Exception as e:
            summary_report.append({"filename": file.name, "status": "Fail", "msg": f"讀取錯誤: {e}"})
            continue

        # 3. 找欄位
        cols = [str(c) for c in df.columns]
        id_col = next((c for c in cols if '身分證' in c), None)
        birth_col = next((c for c in cols if '生日' in c and '民國' in c), None)
        
        stats = {"filename": file.name, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
        if is_encrypted: stats["msg"] += " (含加密)"

        if not id_col or not birth_col:
            summary_report.append({"filename": file.name, "status": "Fail", "msg": "找不到關鍵欄位"})
            continue

        # 4. 檢查與標記
        output = io.BytesIO()
        error_cells = []
        
        id_idx = df.columns.get_loc(id_col)
        birth_idx = df.columns.get_loc(birth_col)

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

        # 5. 寫入 Excel
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            yellow = workbook.add_format({'bg_color': '#FFFF00'})
            
            for r, c in error_cells:
                val = df.iat[r, c]
                if pd.isna(val): val = ""
                worksheet.write(r + 1, c, val, yellow)
            
            worksheet.set_column(0, len(cols)-1, 15)

            # 如果原本有加密，輸出就加密 (維持原始邏輯)
            final_pass = password if (is_encrypted or password) else None
            if final_pass:
                workbook.set_encryption(final_pass)

        processed_files.append((f"已檢查_{file.name}", output.getvalue()))
        summary_report.append(stats)
        progress_bar.progress((i + 1) / len(uploaded_files))
        
    return processed_files, summary_report

def run_encryptor(uploaded_files, new_password):
    """ 分頁 2: 單純加密功能 """
    processed_files = []
    progress_bar = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        try:
            content = file.read()
            # 嘗試直接讀取 (假設使用者上傳的是乾淨的無鎖檔案)
            # 如果是加密檔，這裡會報錯，我們會捕捉它
            df = pd.read_excel(io.BytesIO(content))
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                # 設定密碼
                workbook.set_encryption(new_password)
            
            processed_files.append((f"加密_{file.name}", output.getvalue()))
            
        except Exception as e:
            # 這裡捕捉錯誤 (例如上傳了加密檔但想重新加密)
            st.error(f"❌ {file.name} 失敗: {e} (若檔案原本有加密，請先解鎖再上傳)")
        
        progress_bar.progress((i + 1) / len(uploaded_files))
        
    return processed_files

# ================= 3. 主程式介面 =================

st.set_page_config(page_title="投保名單工具箱", page_icon="🧰")
st.title("🧰 科普列車 - 投保名單工具箱")

tab1, tab2 = st.tabs(["🔍 1. 檢查名單", "🔒 2. 批次加密"])

# --- 分頁 1: 檢查 (原始版本) ---
with tab1:
    st.header("名單檢查工具")
    st.info("此分頁功能：解鎖加密檔 -> 檢查格式 -> 標記黃底 -> (若有密碼則加密回存)。")
    
    check_pass = st.text_input("輸入解鎖密碼 (若檔案無加密可留空)", type="password", key="p1")
    check_files = st.file_uploader("上傳 Excel", type=['xlsx'], accept_multiple_files=True, key="u1")
    
    if check_files and st.button("🚀 開始檢查", key="b1"):
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
                        report_str += f"{item['filename']}: 未滿15歲:{item['under_15']}, 成人:{item['adult']}, 錯誤:{item['errors']}\n"
                    else:
                        report_str += f"{item['filename']}: {item['msg']}\n"
                zf.writestr("報告.txt", report_str)
                
            st.download_button("📦 下載檢查結果 (ZIP)", zip_buffer.getvalue(), "檢查結果.zip", "application/zip")

# --- 分頁 2: 加密 (新功能) ---
with tab2:
    st.header("Excel 批次加密")
    st.info("將無密碼的 Excel 檔加上密碼。")
    
    enc_pass = st.text_input("設定新密碼 (必填)", type="password", key="p2")
    enc_files = st.file_uploader("上傳要加密的 Excel (需無密碼)", type=['xlsx'], accept_multiple_files=True, key="u2")
    
    if enc_files:
        if not enc_pass:
            st.warning("請輸入要設定的密碼！")
        else:
            if st.button("🔒 開始加密", key="b2"):
                enc_results = run_encryptor(enc_files, enc_pass)
                
                if enc_results:
                    st.success(f"成功加密 {len(enc_results)} 個檔案")
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for fname, data in enc_results:
                            zf.writestr(fname, data)
                    
                    st.download_button("📦 下載已加密檔案 (ZIP)", zip_buffer.getvalue(), "已加密.zip", "application/zip")
