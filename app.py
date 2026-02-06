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
    st.error("🛑 缺少必要套件，請檢查 requirements.txt")
    st.stop()

# ================= 1. 核心邏輯區 =================
REF_DATE = datetime(2025, 10, 20)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

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
    """ 安全開啟 Excel (支援加密) """
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
    """ 分頁 1 邏輯：只檢查，不加密回存 """
    wb = open_excel_safe(content, password)
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    ws = wb.active
    col_idx_map = {}
    
    # 找表頭
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
    
    # 開始檢查 (跳過表頭)
    start_row = 2 
    for row in ws.iter_rows(min_row=start_row):
        # 檢查生日
        if xl_birth and xl_birth - 1 < len(row):
            cell = row[xl_birth - 1]
            dt = parse_roc_birthday(cell.value)
            if dt is None:
                cell.fill = YELLOW_FILL
                stats["errors"] += 1
            else:
                age = calculate_age(dt)
                if 0 <= age < 15: stats["under_15"] += 1
                elif age >= 15: stats["adult"] += 1

        # 檢查身分證
        if xl_id and xl_id - 1 < len(row):
            cell = row[xl_id - 1]
            val = str(cell.value).strip() if cell.value else ""
            if not val or val == 'None' or len(val) != 10:
                cell.fill = YELLOW_FILL
                stats["errors"] += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, stats

# ================= 2. 分頁功能 =================

def run_checker(files, pwd):
    processed = []
    report = []
    bar = st.progress(0)
    
    for i, f in enumerate(files):
        data, stats = process_file_logic(f.name, f.read(), pwd)
        report.append(stats)
        if data:
            # 這裡回傳的是 openpyxl 存的檔，絕對沒有密碼
            processed.append((f"已檢查_{f.name}", data.getvalue()))
        bar.progress((i + 1) / len(files))
    return processed, report

def run_encryptor_native(files, pwd):
    """ 使用 xlsxwriter 原生寫入，避開 pandas 引擎衝突 """
    processed = []
    bar = st.progress(0)
    
    for i, f in enumerate(files):
        try:
            content = f.read()
            # 讀取資料
            try:
                df = pd.read_excel(io.BytesIO(content))
            except:
                st.error(f"❌ {f.name}: 讀取失敗，請確認檔案無密碼。")
                continue
            
            # 使用原生 xlsxwriter 寫入加密
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet()
            
            # 寫入資料
            header = df.columns.values
            for c, val in enumerate(header):
                worksheet.write(0, c, str(val))
            
            data = df.fillna("").values
            for r, row in enumerate(data):
                for c, val in enumerate(row):
                    worksheet.write(r + 1, c, val)
            
            # 設定密碼 (這是導致錯誤的關鍵，原生寫法才穩)
            workbook.set_encryption(pwd)
            workbook.close()
            
            output.seek(0)
            processed.append((f"加密_{f.name}", output.getvalue()))
            
        except Exception as e:
            st.error(f"❌ {f.name} 加密失敗: {e}")
        bar.progress((i + 1) / len(files))
    return processed

# ================= 3. 主介面 =================

st.set_page_config(page_title="投保工具箱 V3.0", page_icon="🧰")
st.title("🧰 科普列車 - 投保工具箱 V3.0")

tab1, tab2 = st.tabs(["🔍 1. 檢查名單", "🔒 2. 批次加密"])

with tab1:
    st.header("名單檢查")
    st.info("此頁面檢查後下載的檔案為【無密碼】。確認內容無誤後，請到分頁 2 進行加密。")
    pwd = st.text_input("輸入解鎖密碼 (若檔案無加密可留空)", type="password", key="p1")
    files1 = st.file_uploader("上傳 Excel", type=['xlsx'], accept_multiple_files=True, key="u1")
    
    if files1 and st.button("🚀 開始檢查", key="b1"):
        res, rep = run_checker(files1, pwd)
        if rep: st.dataframe(pd.DataFrame(rep))
        if res:
            z = io.BytesIO()
            with zipfile.ZipFile(z, "w") as zf:
                for n, d in res: zf.writestr(n, d)
                txt = "檢查報告\n" + "\n".join([f"{r['filename']}: {r['msg']}" for r in rep])
                zf.writestr("report.txt", txt)
            st.download_button("📦 下載檢查結果 (ZIP)", z.getvalue(), "檢查結果.zip", "application/zip")

with tab2:
    st.header("批次加密")
    st.warning("請上傳【無密碼】的檔案 (例如從分頁 1 下載的檔案)。")
    new_pwd = st.text_input("設定新密碼", type="password", key="p2")
    files2 = st.file_uploader("上傳要加密的檔案", type=['xlsx'], accept_multiple_files=True, key="u2")
    
    if files2 and new_pwd:
        if st.button("🔒 開始加密", key="b2"):
            res = run_encryptor_native(files2, new_pwd)
            if res:
                st.success(f"成功加密 {len(res)} 個檔案")
                z = io.BytesIO()
                with zipfile.ZipFile(z, "w") as zf:
                    for n, d in res: zf.writestr(n, d)
                st.download_button("📦 下載加密檔案 (ZIP)", z.getvalue(), "已加密.zip", "application/zip")
    elif files2 and not new_pwd:
        st.warning("請輸入密碼！")
