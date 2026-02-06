import streamlit as st
import pandas as pd
import io
import msoffcrypto
from datetime import datetime
import zipfile
import xlsxwriter # 確保有安裝 XlsxWriter

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
    """ 解密檔案，回傳 (BytesIO, 是否原本有加密) """
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

def process_single_file(filename, content, password):
    """ 核心處理邏輯 """
    # 1. 讀取與解密
    decrypted_stream, is_encrypted = get_decrypted_stream(content, password)
    
    if decrypted_stream is None:
        return None, {"filename": filename, "status": "Fail", "msg": "無法開啟 (密碼錯誤或格式不支援)"}

    # 2. 用 Pandas 讀取資料
    try:
        # 自動尋找表頭 (讀前30列判斷)
        preview = pd.read_excel(decrypted_stream, nrows=30, header=None)
        decrypted_stream.seek(0)
        
        header_idx = 0
        for idx, row in preview.iterrows():
            row_str = row.astype(str).values
            if any('身分證' in s for s in row_str) and any('生日' in s for s in row_str):
                header_idx = idx
                break
        
        df = pd.read_excel(decrypted_stream, header=header_idx)
    except Exception as e:
        return None, {"filename": filename, "status": "Fail", "msg": f"讀取失敗: {str(e)}"}

    # 3. 尋找關鍵欄位
    cols = df.columns.tolist()
    id_col_name = next((c for c in cols if '身分證' in str(c)), None)
    birth_col_name = next((c for c in cols if '生日' in str(c) and '民國' in str(c)), None)

    stats = {"filename": filename, "under_15": 0, "adult": 0, "errors": 0, "status": "Success", "msg": "OK"}
    if is_encrypted: stats["msg"] += " (已重新加密)"

    if not id_col_name or not birth_col_name:
