import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
import zipfile
from io import BytesIO
from openpyxl import load_workbook

# Tải dữ liệu mặc định từ GitHub (file2, file3, file4)
@st.cache_data
def load_default_data():
    urls = {
        'file2': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        'file3': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        'file4': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    }
    data = {}
    for key, url in urls.items():
        resp = requests.get(url)
        resp.raise_for_status()
        data[key] = pd.read_excel(BytesIO(resp.content), engine='openpyxl')
    return data['file2'], data['file3'], data['file4']

file2, file3, file4 = load_default_data()

# Chuẩn hóa text
def remove_diacritics(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def normalize_text(s: str) -> str:
    s = str(s)
    s = remove_diacritics(s).lower()
    return re.sub(r'\s+', '', s)

# Chuẩn hóa hoạt chất, hàm lượng, nhóm
def normalize_active(name: str) -> str:
    return re.sub(r'\s+', ' ', re.sub(r'\(.*?\)', '', str(name))).strip().lower()

def normalize_concentration(conc: str) -> str:
    s = str(conc).lower().replace(',', '.')
    parts = [p.strip() for p in re.split(r'[;,]', s) if p.strip()]
    parts = [p for p in parts if re.search(r'\d', p)]
    if len(parts) >= 2 and re.search(r'(mg|mcg|g|%)', parts[0]) and 'ml' in parts[-1]:
        return parts[0].replace(' ', '') + '/' + parts[-1].replace(' ', '')
    return ''.join([p.replace(' ', '') for p in parts])

def normalize_group(grp: str) -> str:
    return re.sub(r'\D', '', str(grp)).strip()

# Sidebar chọn chức năng
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu", "Phân Tích Danh Mục Thầu", "Phân Tích Danh Mục Trúng Thầu", "Đề Xuất Hướng Triển Khai"
])

# 1. Lọc Danh Mục Thầu
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3_temp = file3.copy()
    for col in ['Miền', 'Vùng', 'Tỉnh', 'Bệnh viện/SYT']:
        options = ['(Tất cả)'] + sorted(df3_temp[col].dropna().unique())
        sel = st.selectbox(f"Chọn {col}", options)
        if sel != '(Tất cả)': df3_temp = df3_temp[df3_temp[col] == sel]
    st.session_state['file3_temp'] = df3_temp.copy()

    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        xls = pd.ExcelFile(uploaded, engine='openpyxl')
        sheet = max(xls.sheet_names, key=lambda s: pd.read_excel(uploaded, sheet_name=s, nrows=5, header=None, engine='openpyxl').shape[1])
        try:
            raw = pd.read_excel(uploaded, sheet_name=sheet, header=None, engine='openpyxl')
        except Exception:
            uploaded.seek(0)
            raw_data = uploaded.read()
            zf = zipfile.ZipFile(BytesIO(raw_data), 'r')
            cleaned = BytesIO()
            with zipfile.ZipFile(cleaned, 'w') as w:
                for item in zf.infolist():
                    data = zf.read(item.filename)
                    if item.filename.startswith('xl/worksheets/') or item.filename == 'xl/styles.xml':
                        data = re.sub(b' errorType="[^"]+"', b'', data)
                        data = re.sub(b' errorStyle="[^"]+"', b'', data)
                        data = re.sub(b'<cellStyleXfs.*?</cellStyleXfs>', b'', data, flags=re.DOTALL)
                        data = re.sub(b'<dataValidations.*?</dataValidations>', b'', data, flags=re.DOTALL)
                    w.writestr(item.filename, data)
            cleaned.seek(0)
            wb2 = load_workbook(cleaned, read_only=True, data_only=True)
            ws2 = wb2[sheet]
            rows = list(ws2.iter_rows(values_only=True))
            raw = pd.DataFrame(rows)

        header_idx_auto = None
        scores = []
        for i in range(min(10, len(raw))):
            text = normalize_text(' '.join(raw.iloc[i].fillna('').astype(str).tolist()))
            sc = sum(kw in text for kw in ['tenhoatchat','soluong','nhomthuoc','nongdo'])
            scores.append((i, sc))
            if 'tenhoatchat' in text and 'soluong' in text:
                header_idx_auto = i
                break
        if header_idx_auto is None:
            idx, sc = max(scores, key=lambda x: x[1])
            header_idx_auto = idx if sc>0 else 0
            st.warning(f"Đề xuất dòng tiêu đề: {header_idx_auto+1}")
        st.subheader("🔎 Xem 10 dòng đầu để chọn header (start từ 1)")
        st.dataframe(raw.head(10))
        header_idx = st.number_input("Chọn dòng header (1-10):", 1, min(10, raw.shape[0]), value=header_idx_auto+1) - 1

        header = raw.iloc[header_idx].tolist()
        df_body = raw.iloc[header_idx+1:].copy()
        df_body.columns = header
        df_body = df_body.dropna(subset=header, how='all')
        df_body['_orig_idx'] = df_body.index
        df_body.reset_index(drop=True, inplace=True)

        st.success("✅ Dữ liệu đã được tải và xử lý thành công")
        st.dataframe(df_body.head())
