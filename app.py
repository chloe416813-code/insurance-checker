import streamlit as st
import pandas as pd
import io
import msoffcrypto
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
import zipfile
import xlsxwriter

# ================= 0. 系統環境檢查 =================
try:
    import openpyxl
    import msoffcrypto
    import xlsxwriter
except ImportError:
    st.error("🛑 缺少必要套件")
    st.stop()

# ================= 1. 核心邏輯區 (檢查功能) =================
REF_DATE = datetime(2025, 10, 20)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

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

def open_excel_with_password(file_content, password):
    """ 嘗試開啟 Excel (支援加密與非加密) """
    file_stream = io.BytesIO(file_content)
    # 1. 先嘗試直接開啟
    try:
        wb = openpyxl.load_workbook(file_stream)
        return wb
    except:
        file_stream.seek(0)
    # 2. 嘗試用密碼解鎖
    if password:
        try:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            wb = openpyxl.load_workbook(decrypted)
            return wb
        except:
            return None
    return None

def process_single_file_logic(filename, content, password):
    """ 檢查邏輯 (保留您原始程式碼結構) """
    wb = open_excel_with_password(content, password)

    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟(密碼錯誤或格式不支援)"}

    ws = wb.active
    col_idx_map = {}
    
    # 找表頭 (稍微增強避免空行)
    header_found = False
    for row in ws.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if cell.value: col_idx_map[str(cell.value)] = cell.column
        if '身分證' in col_idx_map or any('身分證' in str(k) for k in col_idx_map.keys()):
            header_found = True
            break
            
    if not header_found:
         col_idx_map = {}
         for row in ws.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value: col_idx_map[str(cell.value)] = cell.column

    id_key = next((k for k in col_idx_map.keys() if '身分證' in k), None)
    birth_key = next((k for k in col_idx_map.keys() if '生日' in k and '民國' in k), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}

    if not id_key or not birth_key:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位"}

    xl_birth_col = col_idx_map[birth_key]
    xl_id_col = col_idx_map[id_key]
    
    # 判斷資料起始列
    start_row = 2 
    for row in ws.iter_rows(min_row=start_row):
        # 1. 檢查生日
        if xl_birth_col and xl_birth_col - 1 < len(row):
            cell_birth = row[xl_birth_col - 1]
            birth_dt = parse_roc_birthday(cell_birth.value)

            if birth_dt is None:
                cell_birth.fill = YELLOW_FILL
                stats["errors"] += 1
            else:
                age = calculate_age(birth_dt)
                if 0 <= age < 15: stats["under_15"] += 1
                elif age >= 15: stats["adult"] += 1

        # 2. 檢查身分證
        if xl_id_col and xl_id_col - 1 < len(row):
            cell_id = row[xl_id_col - 1]
            val_id = str(cell_id.value).strip() if cell_id.value else ""
            if not val_id or val_id == 'None' or len(val_id) != 10:
                cell_id.fill = YELLOW_FILL
                stats["errors"] += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, stats

# ================= 2. 分頁功能實作 =================

def run_checker_tab(uploaded_files, password):
    processed_files = []
    summary_report = []
    progress_bar = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        content = file.read()
        processed_data, stats = process_single_file_logic(file.name, content, password)
        summary_report.append(stats)
        if processed_data:
            processed_files.append((f"已檢查_{file.name}", processed_data.getvalue()))
        progress_bar.progress((i + 1) / len(uploaded_files))
        
    return processed_files, summary_report

def run_encryptor_tab(uploaded_files, new_password):
    """ 分頁 2: 批次加密 (強力清洗版) """
    processed_files = []
    progress_bar = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        try:
            content = file.read()
            file_stream = io.BytesIO(content)
            
            # 1. 檢查檔案是否已加密
            is_already_encrypted = False
            try:
                office_file = msoffcrypto.OfficeFile(file_stream)
                if office_file.is_encrypted():
                    is_already_encrypted = True
            except:
                pass # 不是 Office 檔案或沒加密
            
            if is_already_encrypted:
                st.error(f"❌ {file.name}: 檔案原本就有密碼！請先解鎖成無密碼檔案後再上傳。")
                continue

            # 2. 讀取數據 (清洗數據)
            # 使用 openpyxl 讀取值，不讀取樣式，避免格式干擾
            file_stream.seek(0)
            wb_in = openpyxl.load_workbook(file_stream, data_only=True)
            ws_in = wb_in.active
            
            # 轉成 DataFrame
            data = ws_in.values
            cols = next(data)
            df = pd.DataFrame(data, columns=cols)
            
            # 3. 寫入全新的加密檔案
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                workbook.set_encryption(new_password)  # 設定密碼
            
            processed_files.append((f"加密_{file.name}", output.getvalue()))
            
        except Exception as e:
            st.error(f"❌ {file.name} 加密失敗: {str(e)}")
            
        progress_bar.progress((i + 1) / len(uploaded_files))
        
    return processed_files

# ================= 3. 主程式介面 =================

st.set_page_config(page_title="投保名單工具箱", page_icon="🧰")
st.title("🧰 科普列車 - 投保名單工具箱")

tab1, tab2 = st.tabs(["🔍 1. 檢查名單", "🔒 2. 批次加密"])

# --- 分頁 1: 檢查 ---
with tab1:
    st.header("名單檢查工具")
    st.info("功能：讀取 Excel (支援加密) -> 檢查並標記黃底 -> 輸出 **無密碼** 檔案。")
    st.caption("建議流程：在此頁檢查並下載無密碼檔 -> 確認內容 -> 到分頁 2 進行加密。")
    
    check_pass = st.text_input("輸入解鎖密碼 (若檔案無加密可留空)", type="password", key="p1")
    check_files = st.file_uploader("上傳 Excel", type=['xlsx'], accept_multiple_files=True, key="u1")
    
    if check_files and st.button("🚀 開始檢查", key="b1"):
        results, report = run_checker_tab(check_files, check_pass)
        
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
                zf.writestr("檢查報告.txt", report_str)
                
            st.download_button("📦 下載檢查結果 (ZIP)", zip_buffer.getvalue(), "檢查結果.zip", "application/zip")

# --- 分頁 2: 加密 ---
with tab2:
    st.header("Excel 批次加密")
    st.warning("⚠️ 請注意：此處僅接受 **無密碼** 的 Excel 檔案 (例如剛從分頁 1 下載的檔案)。")
    
    enc_pass = st.text_input("設定新密碼 (必填)", type="password", key="p2")
    enc_files = st.file_uploader("上傳要加密的 Excel (需無密碼)", type=['xlsx'], accept_multiple_files=True, key="u2")
    
    if enc_files:
        if not enc_pass:
            st.warning("請輸入要設定的密碼！")
        else:
            if st.button("🔒 開始加密", key="b2"):
                enc_results = run_encryptor_tab(enc_files, enc_pass)
                
                if enc_results:
                    st.success(f"成功加密 {len(enc_results)} 個檔案")
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        for fname, data in enc_results:
                            zf.writestr(fname, data)
                    
                    st.download_button("📦 下載加密檔案 (ZIP)", zip_buffer.getvalue(), "已加密檔案.zip", "application/zip")
