import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
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
        content = requests.get(url).content
        data[key] = pd.read_excel(BytesIO(content), engine='openpyxl')
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
        if sel != '(Tất cả)':
            df3_temp = df3_temp[df3_temp[col] == sel]
    st.session_state['file3_temp'] = df3_temp.copy()

    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        # built-in sheet detection
        xls = pd.ExcelFile(uploaded, engine='openpyxl')
        sheet = max(xls.sheet_names, key=lambda s: pd.read_excel(uploaded, sheet_name=s, nrows=5, header=None, engine='openpyxl').shape[1])
        # cố gắng đọc với pandas, nếu lỗi thì fallback openpyxl trực tiếp
        try:
            raw = pd.read_excel(uploaded, sheet_name=sheet, header=None, engine='openpyxl')
        except Exception:
            wb = load_workbook(uploaded, read_only=True, data_only=True)
            ws = wb[sheet]
            data = list(ws.values)
            raw = pd.DataFrame(data)

        # Tìm header row
        header_idx = None
        scores = []
        for i in range(min(10, raw.shape[0])):
            text = normalize_text(' '.join(raw.iloc[i].fillna('').astype(str).tolist()))
            sc = sum(kw in text for kw in ['tenhoatchat','soluong','nhomthuoc','nongdo'])
            scores.append((i, sc))
            if 'tenhoatchat' in text and 'soluong' in text:
                header_idx = i
                break
        if header_idx is None:
            idx, sc = max(scores, key=lambda x: x[1])
            if sc > 0:
                header_idx = idx
                st.warning(f"Tự động chọn dòng tiêu đề tại dòng {idx+1}")
            else:
                st.error("❌ Không xác định được header.")
                st.stop()

        # Gán header và body
        header = raw.iloc[header_idx].tolist()
        df_body = raw.iloc[header_idx+1:].copy()
        df_body.columns = header
        df_body = df_body.dropna(subset=header, how='all')
        df_body['_orig_idx'] = df_body.index
        df_body = df_body.reset_index(drop=True)

        # Chuẩn hóa cột
        col_map_body = {}
        for c in df_body.columns:
            n = normalize_text(c)
            if 'tenhoatchat' in n or 'tenthanhphan' in n:
                col_map_body[c] = 'Tên hoạt chất'
            elif 'nongdo' in n or 'hamluong' in n:
                col_map_body[c] = 'Nồng độ/hàm lượng'
            elif 'nhom' in n and 'thuoc' in n:
                col_map_body[c] = 'Nhóm thuốc'
            elif 'soluong' in n:
                col_map_body[c] = 'Số lượng'
            elif 'duongdung' in n or 'duong' in n:
                col_map_body[c] = 'Đường dùng'
            elif 'gia' in n:
                col_map_body[c] = 'Giá kế hoạch'
        df_body.rename(columns=col_map_body, inplace=True)

        # Chuẩn hóa file2
        df2 = file2.copy()
        col_map2 = {}
        for c in df2.columns:
            n = normalize_text(c)
            if 'tenhoatchat' in n:
                col_map2[c] = 'Tên hoạt chất'
            elif 'nongdo' in n or 'hamluong' in n:
                col_map2[c] = 'Nồng độ/hàm lượng'
            elif 'nhom' in n and 'thuoc' in n:
                col_map2[c] = 'Nhóm thuốc'
            elif 'tensanpham' in n:
                col_map2[c] = 'Tên sản phẩm'
        df2.rename(columns=col_map2, inplace=True)

        # Thêm field chuẩn hóa để merge
        for df_ in [df_body, df2]:
            df_['active_norm'] = df_['Tên hoạt chất'].apply(normalize_active)
            df_['conc_norm'] = df_['Nồng độ/hàm lượng'].apply(normalize_concentration)
            df_['group_norm'] = df_['Nhóm thuốc'].apply(normalize_group)

        # Merge và loại duplicate
        merged = pd.merge(df_body, df2, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
        merged = merged.drop_duplicates(subset=['_orig_idx'], keep='first')
        hosp = df3_temp[['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
        merged = pd.merge(merged, hosp, on='Tên sản phẩm', how='left')

        export_df = merged.drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'])
        display_df = merged[merged['_merge']=='both'].drop(columns=['active_norm','conc_norm','group_norm','_merge','_orig_idx'])
        st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
        st.dataframe(display_df)
        st.session_state['filtered_export'] = export_df.copy()
        st.session_state['filtered_display'] = display_df.copy()

        # Tra cứu
        kw = st.text_input("🔍 Tra cứu hoạt chất:")
        if kw:
            df_search = display_df[display_df['Tên hoạt chất'].str.contains(kw, case=False, na=False)]
            st.dataframe(df_search)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False, sheet_name='KetQuaLoc')
        st.download_button('⬇️ Tải File', data=buf.getvalue(), file_name='Ketqua_loc_all.xlsx')

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu (Số liệu)")
    if 'filtered_display' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df = st.session_state['filtered_display'].copy()
        df['Số lượng'] = pd.to_numeric(df['Số lượng'], errors='coerce').fillna(0)
        df['Giá kế hoạch'] = pd.to_numeric(df.get('Giá kế hoạch',0), errors='coerce').fillna(0)
        df['Trị giá'] = df['Số lượng'] * df['Giá kế hoạch']
        def fmt(x):
            if x>=1e9: return f"{x/1e9:.2f} tỷ"
            if x>=1e6: return f"{x/1e6:.2f} triệu"
            if x>=1e3: return f"{x/1e3:.2f} nghìn"
            return str(int(x))
        groups = file4['Nhóm điều trị'].dropna().unique()
        sel_group = st.selectbox("Chọn Nhóm điều trị", ['(Tất cả)'] + list(groups))
        if sel_group != '(Tất cả)':
            acts = file4[file4['Nhóm điều trị']==sel_group]['Tên hoạt chất']
            df = df[df['Tên hoạt chất'].isin(acts)]
        val_act = df.groupby('Tên hoạt chất')['Trị giá'].sum().reset_index().sort_values('Trị giá', ascending=False)
        val_act['Trị giá'] = val_act['Trị giá'].apply(fmt)
        qty_act = df.groupby('Tên hoạt chất')['Số lượng'].sum().reset_index().sort_values('Số lượng', ascending=False)
        qty_act['Số lượng'] = qty_act['Số lượng'].apply(fmt)
        st.subheader('Tổng Trị giá theo Hoạt chất')
        st.table(val_act)
        st.subheader('Tổng Số lượng theo Hoạt chất')
        st.table(qty_act)
        st.subheader('Top 10 Hoạt chất theo Đường dùng & Nhóm điều trị (Số lượng)')
        for route in ['tiêm','uống']:
            sub = df[df['Đường dùng'].str.contains(route, case=False, na=False)]
            top = sub.groupby('Tên hoạt chất')['Số lượng'].sum().nlargest(10).reset_index()
            top['Số lượng'] = top['Số lượng'].apply(fmt)
            st.markdown(f"**{route.capitalize()} - Top 10 theo Số lượng**")
            st.table(top)
        total_sp = df['Tên sản phẩm'].nunique()
        cust_df = df.groupby('Tên Khách hàng phụ trách triển khai').agg(
            SL=('Số lượng','sum'), TG=('Trị giá','sum'), SP=('Tên sản phẩm', pd.Series.nunique)
        ).reset_index()
        cust_df['Tỷ lệ SP'] = (cust_df['SP'] / total_sp * 100).round(2).astype(str) + '%'
        cust_df['SL'] = cust_df['SL'].apply(fmt)
        cust_df['TG'] = cust_df['TG'].apply(fmt)
        st.subheader('Phân tích theo Khách hàng phụ trách')
        st.table(cust_df)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    pass

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered_export' not in st.session_state or 'file3_temp' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df_f = st.session_state['filtered_export'].copy()
        df3t = st.session_state['file3_temp'].copy()
        df3t = df3t[~df3t['Địa bàn'].str.contains('Tạm ngưng triển khai|ko có địa bàn', case=False, na=False)]
        df_qty = df_f.groupby('Tên sản phẩm')['Số lượng'].sum().rename('SL_trúng').reset_index()
        df_sug = pd.merge(df3t, df_qty, on='Tên sản phẩm', how='left').fillna({'SL_trúng':0})
        df_sug = pd.merge(df_sug, file4[['Tên hoạt chất','Nhóm điều trị']], on='Tên hoạt chất', how='left')
        df_sug['Số lượng đề xuất'] = (df_sug['SL_trúng'] * 1.5).apply(np.ceil).astype(int)
        df_sug['Lý do'] = df_sug.apply(
            lambda r: f"Nhóm {r['Nhóm điều trị']} thường sử dụng các hoạt chất tương ứng; sản phẩm chúng ta thế hệ mới, hiệu quả tốt hơn.",
            axis=1
        )
        st.subheader('File 3 tạm & Đề xuất')
        st.dataframe(df_sug)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            df_sug.to_excel(writer, index=False, sheet_name='DeXuat')
        st.download_button('⬇️ Tải File Đề Xuất', data=buf.getvalue(), file_name='DeXuat_Thuoc.xlsx')
