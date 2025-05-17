import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
import zipfile
from io import BytesIO
from openpyxl import load_workbook
from datetime import datetime

# ============ Định nghĩa hàm safe_str ở ngoài ============
def safe_str(x):
    try:
        if pd.isnull(x):
            return ''
        return str(x)
    except Exception:
        return ''

# ... (các hàm load file, chuẩn hóa khác như cũ)

# === Load default data from GitHub ===
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

# ... (các hàm chuẩn hóa text, group v.v. như cũ)
def remove_diacritics(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def normalize_text(s: str) -> str:
    s = str(s)
    s = remove_diacritics(s).lower()
    return re.sub(r'\s+', '', s)

def normalize_active(name: str) -> str:
    return re.sub(r'\s+', ' ', re.sub(r'\(.*?\)', '', str(name))).strip().lower()

def normalize_concentration(conc: str) -> str:
    s = str(conc).lower().replace(',', '.')
    parts = [p.strip() for p in re.split(r'[;,]', s) if p.strip()]
    parts = [p for p in parts if re.search(r'\d', p)]
    if len(parts) >= 2 and re.search(r'(mg|mcg|g|%)', parts[0]) and 'ml' in parts[-1]:
        return parts[0].replace(' ', '') + '/' + parts[-1].replace(' ', '')
    return ''.join(p.replace(' ', '') for p in parts)

def normalize_group(grp: str) -> str:
    return re.sub(r'\D', '', str(grp)).strip()

# ============ Hàm xử lý file upload =============
def process_uploaded(uploaded, df3_temp):
    # Tìm sheet nhiều cột nhất
    xls = pd.ExcelFile(uploaded, engine='openpyxl')
    sheet = max(xls.sheet_names, key=lambda s: pd.read_excel(uploaded, sheet_name=s, nrows=5, header=None, engine='openpyxl').shape[1])
    raw = pd.read_excel(uploaded, sheet_name=sheet, header=None, engine='openpyxl')
    # Tìm header dòng nào
    header_idx = None
    scores = []
    for i in range(min(10, len(raw))):
        text = normalize_text(' '.join(raw.iloc[i].fillna('').astype(str).tolist()))
        sc = sum(kw in text for kw in ['tenhoatchat','soluong','nhom','nongdo','thanhphan','hamluong'])
        scores.append((i, sc))
        if 'tenhoatchat' in text and 'soluong' in text:
            header_idx = i
            break
    if header_idx is None:
        idx, sc = max(scores, key=lambda x: x[1])
        header_idx = idx if sc > 0 else 0
    header = raw.iloc[header_idx].fillna('').astype(str).tolist()
    df_body = raw.iloc[header_idx+1:].copy()
    df_body.columns = header
    df_body = df_body.dropna(subset=header, how='all')
    df_body['_orig_idx'] = df_body.index
    df_body.reset_index(drop=True, inplace=True)

    # ============= MAP tên cột mở rộng =============
    col_map = {}
    for c in df_body.columns:
        n = normalize_text(c)
        if ('tenhoatchat' in n) or ('tenthanhphan' in n) or ('hoatchat' in n and 'ten' in n) or ('thanhphan' in n):
            col_map[c] = 'Tên hoạt chất'
        elif ('nongdo' in n) or ('hamluong' in n) or ('nongdo' in n and 'hamluong' in n) or ('nong do' in c.lower()) or ('hàm lượng' in c.lower()):
            col_map[c] = 'Nồng độ/hàm lượng'
        elif 'nhom' in n:
            col_map[c] = 'Nhóm thuốc'
        elif 'soluong' in n:
            col_map[c] = 'Số lượng'
        elif ('duongdung' in n) or ('duong' in n):
            col_map[c] = 'Đường dùng'
        elif 'gia' in n:
            col_map[c] = 'Giá kế hoạch'
        elif ('tensanpham' in n) or ('sanpham' in n):
            col_map[c] = 'Tên sản phẩm'
    df_body.rename(columns=col_map, inplace=True)

    # Chuẩn hóa df2 (file2)
    df2_norm = file2.copy()
    col_map2 = {}
    for c in df2_norm.columns:
        n = normalize_text(c)
        if 'tenhoatchat' in n:
            col_map2[c] = 'Tên hoạt chất'
        elif 'nongdo' in n or 'hamluong' in n:
            col_map2[c] = 'Nồng độ/hàm lượng'
        elif 'nhom' in n and 'thuoc' in n:
            col_map2[c] = 'Nhóm thuốc'
        elif 'tensanpham' in n:
            col_map2[c] = 'Tên sản phẩm'
    df2_norm.rename(columns=col_map2, inplace=True)

    # Chuẩn hóa các key merge
    for df_ in (df_body, df2_norm):
        df_['active_norm'] = df_['Tên hoạt chất'].apply(normalize_active)
        df_['conc_norm'] = df_['Nồng độ/hàm lượng'].apply(normalize_concentration)
        df_['group_norm'] = df_['Nhóm thuốc'].apply(normalize_group) if 'Nhóm thuốc' in df_.columns else ''

    # Merge
    merged = pd.merge(df_body, df2_norm, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
    merged.drop_duplicates(subset=['_orig_idx'], keep='first', inplace=True)
    hosp = df3_temp[['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']] if 'Tên sản phẩm' in df3_temp.columns else None
    if hosp is not None:
        merged = pd.merge(merged, hosp, on='Tên sản phẩm', how='left')

    export_df = merged.drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'], errors='ignore')
    display_df = merged[merged['_merge']=='both'].drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'], errors='ignore')
    return display_df, export_df

# ==================== MAIN UI =====================
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai"
])

if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3_temp = file3.copy()
    for col in ['Miền','Vùng','Tỉnh','Bệnh viện/SYT']:
        opts = ['(Tất cả)'] + sorted(df3_temp[col].dropna().unique())
        sel = st.selectbox(f"Chọn {col}", opts)
        if sel != '(Tất cả)':
            df3_temp = df3_temp[df3_temp[col]==sel]
    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        display_df, export_df = process_uploaded(uploaded, df3_temp)
        st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
        # CHỐT: Dùng safe_str cho mọi giá trị, không bao giờ lỗi dataframe
        if display_df is not None and not display_df.empty:
            display_df_fix = display_df.applymap(safe_str)
            st.dataframe(display_df_fix)
        else:
            st.info("Không có dòng nào khớp hoặc dữ liệu rỗng.")

        # Save for analysis
        st.session_state['filtered_df'] = export_df.copy()
        st.session_state['selected_hospital'] = df3_temp['Bệnh viện/SYT'].iloc[0] if 'Bệnh viện/SYT' in df3_temp.columns else ''
        # Download filtered file with custom name
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False, sheet_name='KetQuaLoc')
        today = datetime.now().strftime('%d.%m.%y')
        hospital = st.session_state.get('selected_hospital', '').replace('/', '-')
        filename = f"{today}-KQ Loc Thau - {hospital}.xlsx"
        st.download_button('⬇️ Tải File Kết Quả', data=buf.getvalue(), file_name=filename)
