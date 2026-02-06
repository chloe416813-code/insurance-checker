import streamlit as st
import pandas as pd
import io
import msoffcrypto
from datetime import datetime
import zipfile

# ================= 設定區 =================
REF_DATE = datetime(2025, 10, 20)

# ================= 函式區 =================
def parse_roc_birthday(roc_val):
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
    """ 解密檔案串流，回傳 (BytesIO, 是否原本有加密) """
    file_stream = io.BytesIO(file_content)
    # 1. 嘗試直接讀取
    try:
        pd.read_excel(file_stream, nrows=1)
        file_stream.seek(0)
        return file_stream, False
    except:
        file_stream.seek(0)
    
    # 2. 嘗試解密
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

def find_header_row(file_stream):
    """ 自動尋找表頭所在的列數 """
    # 讀取前 20 列來找關鍵字
    df_preview = pd.read_excel(file_stream, header=None, nrows=20)
    file_stream.seek(0)
    
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).values
        if any('身分證' in str(x) for x in row_str) and any('生日' in str(x) for x in row_str):
            return idx
    return 0 # 預設第一列

def highlight_errors(row, id_col, birth_col):
    """ Pandas Style 用的邏輯函式 """
    styles = [''] * len(row)
    yellow = 'background-color: yellow'
    
    # 檢查生日
    birth_val = row[birth_col]
    birth_dt = parse_roc_birthday(birth_val)
    if birth_dt is None:
        # 找到生日欄位的 index 並標記
        idx = row.index.get_loc(birth_col)
        styles[idx] = yellow
    
    # 檢查身分證
    id_val = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
    if not id_val or id_val == 'nan' or len(id_val) != 10:
        idx = row.index.get_loc(id_col)
        styles[idx] = yellow
        
    return styles

def process_single_file(filename, content, password):
    # 解密並讀取
    decrypted_stream, is_encrypted = get_decrypted_stream(content, password)
    
    if decrypted_stream is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    # 尋找表頭
    header_idx = find_header_row(decrypted_stream)
    
    # 讀取資料
    try:
        df = pd.read_excel(decrypted_stream, header=header_idx)
    except Exception as e:
        return None, {"filename": filename, "status": "Fail", "msg": f"讀取失敗: {str(e)}"}

    # 尋找關鍵欄位
    cols = df.columns.tolist()
    id_col = next((c for c in cols if '身分證' in str(c)), None)
    birth_col = next((c for c in cols if '生日' in str(c) and '民國' in str(c)), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (已重新加密)"

    if not id_col or not birth_col:
        return None, {"filename": filename, "status": "Fail", "msg": "找不到關鍵欄位 (需有'身分證'與'生日(民國)')"}

    # 統計數據 (不影響原始資料，只做計算)
    for index, row in df.iterrows():
        birth_dt = parse_roc_birthday(row[birth_col])
        if birth_dt:
            age = calculate_age(birth_dt)
            if 0 <= age < 15: stats["under_15"] += 1
            elif age >= 15: stats["adult"] += 1
        else:
            stats["errors"] += 1
        
        id_val = str(row[id_col]).strip()
        if not id_val or id_val == 'nan' or len(id_val) != 10:
             # 注意：這裡只算錯誤數，樣式標記交給 Pandas Style
             if not (birth_dt is None): # 避免重複計數
                 stats["errors"] += 1

    # 使用 Pandas Styler 進行標記 (黃底)
    # axis=1 表示逐列處理
    styled_df = df.style.apply(highlight_errors, axis=1, id_col=id_col, birth_col=birth_col)

    # 輸出到 Excel (使用 XlsxWriter 引擎以支援加密)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        styled_df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # 設定欄寬 (稍微美化)
        worksheet.set_column(0, len(cols)-1, 15)

        # 若原本有加密 (或使用者有輸入密碼)，則對新檔案加密
        final_password = password if (is_encrypted or password) else None
        if final_password:
            workbook.set_encryption(final_password)
    
    output.seek(0)
    return output, stats

# ================= 網頁介面 (UI) =================
st.set_page_config(page_title="投保名單檢查工具", page_icon="🚄")

st.title("🚄 科普列車 - 投保名單自動檢查工具")
st.markdown(f"**檢查標準日：{REF_DATE.date()}**")
st.info("功能：統計年齡、檢查格式、標記黃底。輸出之 Excel 將會加密保護 (使用您輸入的密碼)。")

# 側邊欄
with st.sidebar:
    st.header("⚙️ 設定")
    password = st.text_input("檔案密碼", type="password")
    st.caption("1. 若上傳加密檔，請輸入解鎖密碼。\n2. 處理後的檔案也會用此密碼加密。")

# 上傳區
uploaded_files = st.file_uploader("請選擇 Excel 檔案", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 開始檢查", type="primary"):
        progress_bar = st.progress(0)
        processed_files = []
        summary_report = []
        
        for i, file in enumerate(uploaded_files):
            content = file.read()
            processed_data, stats = process_single_file(file.name, content, password)
            
            summary_report.append(stats)
            if processed_data:
                processed_files.append((f"已檢查_{file.name}", processed_data))
            
            progress_bar.progress((i + 1) / len(uploaded_files))

        st.success("檢查完成！統計結果如下：")
        st.dataframe(pd.DataFrame(summary_report))

        if processed_files:
            zip_buffer = io.BytesIO()
            # 使用標準 ZIP (不加密)，但裡面的 Excel 是加密的
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for fname, f_data in processed_files:
                    zf.writestr(fname, f_data.getvalue())
                
                # 報告
                report_str = f"【檢查統計報告 - {datetime.now().strftime('%Y-%m-%d %H:%M')}】\n\n"
                for item in summary_report:
                    report_str += f"📄 {item['filename']}: {item['msg']}\n"
                    if item['status'] == 'Success':
                        report_str += f"   - 未滿15歲: {item['under_15']}\n   - 成人: {item['adult']}\n   - 錯誤數(含生日/ID): {item['errors']}\n"
                    report_str += "-"*30 + "\n"
                zf.writestr("總表統計.txt", report_str)

            st.download_button(
                label="📦 下載檢查結果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="檢查結果打包.zip",
                mime="application/zip"
            )
        else:
            st.error("沒有檔案成功處理，請檢查密碼或格式。")
