import streamlit as st
import pandas as pd
import io
import msoffcrypto
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill
import pyzipper  # 改用這個來做加密壓縮

# ================= 設定區 =================
REF_DATE = datetime(2025, 10, 20)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# ================= 函式區 =================
def parse_roc_birthday(roc_val):
    if roc_val is None: return None
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

def open_excel_with_password(file_content, password):
    """ 嘗試開啟 Excel """
    file_stream = io.BytesIO(file_content)
    
    # 1. 先嘗試直接開啟 (無加密)
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

def process_single_file(filename, content, password):
    wb = open_excel_with_password(content, password)
    
    if wb is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟(密碼錯誤或格式不支援)"}

    ws = wb.active
    
    # 自動尋找欄位
    header_row = None
    col_idx_map = {}
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            if cell.value:
                col_idx_map[str(cell.value)] = cell.column

    id_key = next((k for k in col_idx_map.keys() if '身分證' in k), None)
    birth_key = next((k for k in col_idx_map.keys() if '生日' in k and '民國' in k), None)
    
    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}

    if not id_key or not birth_key:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位"}

    xl_birth_col = col_idx_map[birth_key]
    xl_id_col = col_idx_map[id_key]

    for row in ws.iter_rows(min_row=2):
        # 檢查生日
        if xl_birth_col:
            cell_birth = row[xl_birth_col - 1]
            birth_dt = parse_roc_birthday(cell_birth.value)
            if birth_dt is None:
                cell_birth.fill = YELLOW_FILL
                stats["errors"] += 1
            else:
                age = calculate_age(birth_dt)
                if 0 <= age < 15: stats["under_15"] += 1
                elif age >= 15: stats["adult"] += 1

        # 檢查身分證
        if xl_id_col:
            cell_id = row[xl_id_col - 1]
            val_id = str(cell_id.value).strip() if cell_id.value else ""
            if not val_id or val_id == 'None' or len(val_id) != 10:
                cell_id.fill = YELLOW_FILL
                stats["errors"] += 1

    # 直接存成 BytesIO (不加密 Excel 本體，改為最後加密 ZIP)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, stats

# ================= 網頁介面 (UI) =================
st.set_page_config(page_title="投保名單檢查工具", page_icon="🚄")

st.title("🚄 科普列車 - 投保名單自動檢查工具")
st.markdown(f"**檢查標準日：{REF_DATE.date()}**")
st.info("功能：自動統計年齡、檢查身分證格式、針對錯誤欄位標記黃底。")

# 側邊欄：設定與密碼
with st.sidebar:
    st.header("⚙️ 設定")
    password = st.text_input("檔案密碼 (若無加密可留空)", type="password")
    st.caption("輸入密碼後，系統會用此密碼解鎖 Excel 進行檢查。")
    st.markdown("---")
    st.warning("⚠️ 注意：檢查後的檔案將被打包成「加密 ZIP」。解壓縮密碼與您上方輸入的密碼相同 (若未輸入則無密碼)。")

# 檔案上傳區
uploaded_files = st.file_uploader("請拖曳或選擇 Excel 檔案 (可多選)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始檢查", type="primary"):
        progress_bar = st.progress(0)
        processed_files = []
        summary_report = []
        
        for i, file in enumerate(uploaded_files):
            content = file.read()
            # 重設指標，避免讀取問題
            file.seek(0)
            
            processed_data, stats = process_single_file(file.name, content, password)
            
            summary_report.append(stats)
            if processed_data:
                processed_files.append((f"已檢查_{file.name}", processed_data))
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        st.success("檢查完成！統計結果如下：")
        df_report = pd.DataFrame(summary_report)
        st.dataframe(df_report)

        if processed_files:
            zip_buffer = io.BytesIO()
            
            # 判斷是否要加密 ZIP
            if password:
                # 使用 AES 加密建立 ZIP
                with pyzipper.AESZipFile(zip_buffer, "w", compression=pyzipper.ZIP_LZMA, encryption=pyzipper.WZ_AES) as zf:
                    zf.setpassword(password.encode('utf-8'))
                    
                    # 加入 Excel
                    for fname, f_data in processed_files:
                        zf.writestr(fname, f_data.getvalue())
                    
                    # 加入文字報告
                    report_str = f"【檢查統計報告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}】\n\n"
                    for item in summary_report:
                        report_str += f"📄 {item['filename']}: {item['msg']}\n"
                        if item['status'] == 'Success':
                            report_str += f"   - 未滿15歲: {item['under_15']}\n   - 成人: {item['adult']}\n   - 錯誤數: {item['errors']}\n"
                        report_str += "-"*30 + "\n"
                    zf.writestr("總表統計.txt", report_str)
            else:
                # 一般無密碼 ZIP
                with pyzipper.ZipFile(zip_buffer, "w") as zf:
                    for fname, f_data in processed_files:
                        zf.writestr(fname, f_data.getvalue())
                    
                    report_str = f"【檢查統計報告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}】\n\n"
                    for item in summary_report:
                        report_str += f"📄 {item['filename']}: {item['msg']}\n"
                        if item['status'] == 'Success':
                            report_str += f"   - 未滿15歲: {item['under_15']}\n   - 成人: {item['adult']}\n   - 錯誤數: {item['errors']}\n"
                        report_str += "-"*30 + "\n"
                    zf.writestr("總表統計.txt", report_str)

            st.download_button(
                label="📦 下載檢查結果 (ZIP壓縮檔)",
                data=zip_buffer.getvalue(),
                file_name="檢查結果打包.zip",
                mime="application/zip"
            )
        else:
            st.error("沒有檔案被成功處理，請檢查密碼或檔案格式。")
