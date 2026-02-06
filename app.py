import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import sys

# ================= 0. 系統診斷區 (Debug) =================
# 這段程式碼會幫助我們確認環境是否正常
try:
    import openpyxl
    import msoffcrypto
    import xlsxwriter
except ImportError as e:
    st.error(f"🛑 嚴重錯誤：缺少套件 {e}")
    st.stop()

# ================= 1. 核心邏輯區 =================
REF_DATE = datetime(2025, 10, 20)

def parse_roc_birthday(roc_val):
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

def open_excel_safe(file_content, password):
    file_stream = io.BytesIO(file_content)
    try:
        return openpyxl.load_workbook(file_stream)
    except:
        file_stream.seek(0)
    
    if password:
        try:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(file_stream)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return openpyxl.load_workbook(decrypted)
        except:
            return None
    return None

def process_file_logic(filename, content, password):
    """ 分頁 1: 檢查邏輯 (使用 openpyxl) """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    ws = wb.active
    col_idx_map = {}
    
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

    xl_birth = col_idx_map[birth_key]
    xl_id = col_idx_map[id_key]
    
    # 這裡需要重新定義黃色，因為 openpyxl 版本可能不同
    from openpyxl.styles import PatternFill
    YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    start_row = 2 
    for row in ws.iter_rows(min_row=start_row):
        if xl_birth and xl_birth - 1 < len(row):
            cell = row[xl_birth - 1]
            dt = parse_roc_birthday(cell.value)
            if dt is None:
                cell.fill = YELLOW
                stats["errors"] += 1
            else:
                age = calculate_age(dt)
                if 0 <= age < 15: stats["under_15"] += 1
                elif age >= 15: stats["adult"] += 1

        if xl_id and xl_id - 1 < len(row):
            cell = row[xl_id - 1]
            val = str(cell.value).strip() if cell.value else ""
            if not val or val == 'None' or len(val) != 10:
                cell.fill = YELLOW
                stats["errors"] += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, stats

# ================= 2. 執行函式 =================

def run_checker(files, pwd):
    processed = []
    report = []
    bar = st.progress(0)
    for i, f in enumerate(files):
        data, stats = process_file_logic(f.name, f.read(), pwd)
        report.append(stats)
        if data:
            processed.append((f"已檢查_{f.name}", data.getvalue()))
        bar.progress((i + 1) / len(files))
    return processed, report

def run_encryptor_debug(files, pwd):
    """ 分頁 2: 診斷式加密 """
    processed = []
    bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            content = f.read()
            # 讀取
            try:
                df = pd.read_excel(io.BytesIO(content))
            except:
                st.error(f"❌ {f.name}: 讀取失敗，請確認檔案無密碼。")
                continue
            
            # 寫入
            output = io.BytesIO()
            
            # --- 關鍵診斷點 ---
            # 我們強制使用 xlsxwriter，並在出錯時印出物件類型
            try:
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                worksheet = workbook.add_worksheet()
                
                # 寫資料
                header = df.columns.values
                for c, val in enumerate(header):
                    worksheet.write(0, c, str(val))
                data = df.fillna("").values
                for r, row in enumerate(data):
                    for c, val in enumerate(row):
                        worksheet.write(r + 1, c, val)
                
                # 嘗試加密
                if hasattr(workbook, 'set_encryption'):
                    workbook.set_encryption(pwd)
                else:
                    # 萬一真的發生靈異現象，這裡會抓到
                    raise Exception(f"物件類型錯誤: {type(workbook)}，它沒有 set_encryption 方法")
                
                workbook.close()
                output.seek(0)
                processed.append((f"加密_{f.name}", output.getvalue()))
                
            except Exception as inner_e:
                st.error(f"❌ {f.name} 寫入階段失敗: {inner_e}")
                
        except Exception as e:
            st.error(f"❌ {f.name} 整體失敗: {e}")
        bar.progress((i + 1) / len(files))
    return processed

# ================= 3. 主介面 =================

st.set_page_config(page_title="投保工具箱 V4.0 (診斷版)", page_icon="🛠️")
st.title("🛠️ 投保工具箱 V4.0 (診斷版)")

# 顯示環境資訊 (Debug info)
with st.expander("ℹ️ 系統環境資訊 (若報錯請截圖此處)"):
    st.write(f"XlsxWriter Version: {xlsxwriter.__version__}")
    st.write(f"Python Version: {sys.version}")

tab1, tab2 = st.tabs(["🔍 1. 檢查名單", "🔒 2. 批次加密"])

with tab1:
    st.header("名單檢查")
    st.info("檢查後輸出【無密碼】檔案。")
    pwd = st.text_input("輸入解鎖密碼", type="password", key="p1")
    files1 = st.file_uploader("上傳 Excel", type=['xlsx'], accept_multiple_files=True, key="u1")
    
    if files1 and st.button("🚀 開始檢查", key="b1"):
        res, rep = run_checker(files1, pwd)
        if rep: st.dataframe(pd.DataFrame(rep))
        if res:
            z = io.BytesIO()
            with zipfile.ZipFile(z, "w") as zf:
                for n, d in res: zf.writestr(n, d)
                txt = "\n".join([f"{r['filename']}: {r['msg']}" for r in rep])
                zf.writestr("report.txt", txt)
            st.download_button("📦 下載結果", z.getvalue(), "檢查結果.zip", "application/zip")

with tab2:
    st.header("批次加密")
    st.warning("請上傳無密碼檔案。")
    new_pwd = st.text_input("設定新密碼", type="password", key="p2")
    files2 = st.file_uploader("上傳加密檔案", type=['xlsx'], accept_multiple_files=True, key="u2")
    
    if files2 and new_pwd:
        if st.button("🔒 開始加密", key="b2"):
            res = run_encryptor_debug(files2, new_pwd)
            if res:
                st.success(f"加密成功 {len(res)} 個")
                z = io.BytesIO()
                with zipfile.ZipFile(z, "w") as zf:
                    for n, d in res: zf.writestr(n, d)
                st.download_button("📦 下載加密檔", z.getvalue(), "已加密.zip", "application/zip")
