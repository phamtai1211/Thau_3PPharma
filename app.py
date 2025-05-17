import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
import zipfile
from io import BytesIO
from openpyxl import load_workbook
import plotly.express as px
from datetime import datetime

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

# === Text normalization helpers ===
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

# === Process uploaded Excel files ===
def process_uploaded(uploaded, df3_temp):
    # Determine sheet with most columns
    xls = pd.ExcelFile(uploaded, engine='openpyxl')
    sheet = max(xls.sheet_names, key=lambda s: pd.read_excel(uploaded, sheet_name=s, nrows=5, header=None, engine='openpyxl').shape[1])
    try:
        raw = pd.read_excel(uploaded, sheet_name=sheet, header=None, engine='openpyxl')
    except Exception:
        uploaded.seek(0)
        buf = BytesIO(uploaded.read())
        zf = zipfile.ZipFile(buf, 'r')
        out = BytesIO()
        with zipfile.ZipFile(out, 'w') as w:
            for item in zf.infolist():
                if item.filename.startswith('xl/styles') or item.filename.startswith('xl/theme'):
                    continue
                data = zf.read(item.filename)
                if item.filename.startswith('xl/worksheets/'):
                    data = re.sub(b'<dataValidations.*?</dataValidations>', b'', data, flags=re.DOTALL)
                w.writestr(item.filename, data)
        out.seek(0)
        wb = load_workbook(out, read_only=True, data_only=True)
        ws = wb[sheet]
        raw = pd.DataFrame(list(ws.iter_rows(values_only=True)))

    # Auto-detect header row among first 10
    header_idx = None
    scores = []
    for i in range(min(10, len(raw))):
        text = normalize_text(' '.join(raw.iloc[i].fillna('').astype(str).tolist()))
        sc = sum(kw in text for kw in ['tenhoatchat','soluong','nhomthuoc','nongdo'])
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

    # Map to standard names
    col_map = {}
    for c in df_body.columns:
        n = normalize_text(c)
        if 'tenhoatchat' in n or 'tenthanhphan' in n:
            col_map[c] = 'Tên hoạt chất'
        elif 'nongdo' in n or 'hamluong' in n:
            col_map[c] = 'Nồng độ/hàm lượng'
        elif ('nhom' in n and 'thuoc' in n) or (n.startswith('nhom') and len(n) <= 7):  # thêm dòng này
            col_map[c] = 'Nhóm thuốc'
        elif 'soluong' in n:
            col_map[c] = 'Số lượng'
        elif 'duongdung' in n or 'duong' in n:
            col_map[c] = 'Đường dùng'
        elif 'gia' in n:
            col_map[c] = 'Giá kế hoạch'
    df_body.rename(columns=col_map, inplace=True)

    # Prepare reference df2
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

    # Add merge keys
    for df_ in (df_body, df2_norm):
        df_['active_norm'] = df_['Tên hoạt chất'].apply(normalize_active)
        df_['conc_norm'] = df_['Nồng độ/hàm lượng'].apply(normalize_concentration)
        df_['group_norm'] = df_['Nhóm thuốc'].apply(normalize_group)

    merged = pd.merge(df_body, df2_norm, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
    merged.drop_duplicates(subset=['_orig_idx'], keep='first', inplace=True)

    hosp = df3_temp[['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
    merged = pd.merge(merged, hosp, on='Tên sản phẩm', how='left')

    export_df = merged.drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'])
    display_df = merged[merged['_merge']=='both'].drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'])
    return display_df, export_df

# === Main UI ===
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai"
])

# 1. Lọc Danh Mục Thầu
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
        st.dataframe(display_df.fillna('').astype(str))
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

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if 'filtered_df' not in st.session_state:
        st.info("Vui lòng thực hiện bước 'Lọc Danh Mục Thầu' trước.")
    else:
        df = st.session_state['filtered_df'].copy()
        # Force rename to standard columns if misnamed
        rename_map = {}
        for c in df.columns:
            n = normalize_text(c)
            if 'nhom' in n and 'thuoc' in n:
                rename_map[c] = 'Nhóm thuốc'
            if 'trigia' in n or (n.startswith('tri') and 'gia' in n):
                rename_map[c] = 'Trị giá'
        if rename_map:
            df.rename(columns=rename_map, inplace=True)
        df['Số lượng'] = pd.to_numeric(df.get('Số lượng', 0), errors='coerce').fillna(0)
        df['Giá kế hoạch'] = pd.to_numeric(df.get('Giá kế hoạch', 0), errors='coerce').fillna(0)
        df['Trị giá'] = df['Số lượng'] * df['Giá kế hoạch']
        # Chart 1: Trị giá theo Nhóm thuốc
        grp_val = df.groupby('Nhóm thuốc')['Trị giá'].sum().reset_index().sort_values('Trị giá', False)
        fig1 = px.bar(grp_val, x='Nhóm thuốc', y='Trị giá', title='Trị giá theo Nhóm thuốc')
        st.plotly_chart(fig1, use_container_width=True)
        # Chart 2: Tỷ trọng trị giá theo đường dùng
        df['Loại đường dùng'] = df['Đường dùng'].apply(lambda x: 'Tiêm' if 'tiêm' in str(x).lower() else ('Uống' if 'uống' in str(x).lower() else 'Khác'))
        route_val = df.groupby('Loại đường dùng')['Trị giá'].sum().reset_index()
        fig2 = px.pie(route_val, names='Loại đường dùng', values='Trị giá', title='Tỷ trọng trị giá theo đường dùng')
        st.plotly_chart(fig2, use_container_width=True)
        # Chart 3 & 4: Top 10 hoạt chất theo SL và TG
        top_qty = df.groupby('Tên hoạt chất')['Số lượng'].sum().reset_index().sort_values('Số lượng', False).head(10)
        fig3 = px.bar(top_qty, x='Tên hoạt chất', y='Số lượng', title='Top 10 Hoạt chất (SL)')
        st.plotly_chart(fig3, use_container_width=True)
        top_val = df.groupby('Tên hoạt chất')['Trị giá'].sum().reset_index().sort_values('Trị giá', False).head(10)
        fig4 = px.bar(top_val, x='Tên hoạt chất', y='Trị giá', title='Top 10 Hoạt chất (TG)')
        st.plotly_chart(fig4, use_container_width=True)
        # Chart 5: Trị giá theo Nhóm điều trị
        treat_map = {normalize_active(a): grp for a, grp in zip(file4['Hoạt chất'], file4['Nhóm điều trị'])}
        df['Nhóm điều trị'] = df['Tên hoạt chất'].apply(lambda x: treat_map.get(normalize_active(x), 'Khác'))
        treat_val = df.groupby('Nhóm điều trị')['Trị giá'].sum().reset_index().sort_values('Trị giá', False)
        fig5 = px.bar(treat_val, x='Trị giá', y='Nhóm điều trị', orientation='h', title='Trị giá theo Nhóm điều trị')
        st.plotly_chart(fig5, use_container_width=True)
        sel_grp = st.selectbox('Chọn nhóm để xem Top 10 sản phẩm', treat_val['Nhóm điều trị'].tolist())
        if sel_grp:
            top_prod = df[df['Nhóm điều trị']==sel_grp].groupby('Tên sản phẩm')['Trị giá'].sum().reset_index().sort_values('Trị giá', False).head(10)
            fig6 = px.bar(top_prod, x='Trị giá', y='Tên sản phẩm', orientation='h', title=f'Top 10 sản phẩm - Nhóm {sel_grp}')
            st.plotly_chart(fig6, use_container_width=True)
        # Chart 6: Trị giá theo Khách hàng
        rep_val = df.groupby('Tên Khách hàng phụ trách triển khai')['Trị giá'].sum().reset_index().sort_values('Trị giá', False)
        fig7 = px.bar(rep_val, x='Trị giá', y='Tên Khách hàng phụ trách triển khai', orientation='h', title='Trị giá theo Khách hàng phụ trách')
        st.plotly_chart(fig7, use_container_width=True)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    st.info("Chức năng đang xây dựng...")

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    st.info("Chức năng đang xây dựng...")
