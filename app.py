import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
from io import BytesIO

# Tải dữ liệu mặc định từ GitHub (file2, file3, file4)
@st.cache_data
def load_default_data():
    urls = {
        'file2': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        'file3': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        'file4': "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    }
    data = {}
    for k, url in urls.items():
        content = requests.get(url).content
        data[k] = pd.read_excel(BytesIO(content))
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

# Sidebar
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu", "Phân Tích Danh Mục Thầu", "Phân Tích Danh Mục Trúng Thầu", "Đề Xuất Hướng Triển Khai"
])

# 1. Lọc Danh Mục Thầu
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    # chọn khu vực
    regions = sorted(file3['Miền'].dropna().unique())
    sel_region = st.selectbox("Chọn Miền", ['(Tất cả)'] + regions)
    df3_temp = file3.copy()
    if sel_region != '(Tất cả)': df3_temp = df3_temp[df3_temp['Miền'] == sel_region]
    for col in ['Vùng','Tỉnh','Bệnh viện/SYT']:
        vals = sorted(df3_temp[col].dropna().unique())
        sel = st.selectbox(f"Chọn {col}", ['(Tất cả)'] + vals)
        if sel != '(Tất cả)': df3_temp = df3_temp[df3_temp[col] == sel]
    st.session_state['file3_temp'] = df3_temp.copy()

    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        # đọc sheet có nhiều cột nhất
        xls = pd.ExcelFile(uploaded)
        sheet = max(xls.sheet_names, key=lambda s: pd.read_excel(uploaded, sheet_name=s, nrows=5, header=None).shape[1])
        raw = pd.read_excel(uploaded, sheet_name=sheet, header=None)
        # tìm header row
        header_idx = None
        scores = []
        for i in range(min(10, raw.shape[0])):
            row = raw.iloc[i].fillna('').astype(str).tolist()
            text = normalize_text(' '.join(row))
            score = sum([1 for kw in ['tenhoatchat','soluong','nhomthuoc','nongdo'] if kw in text])
            scores.append((i, score))
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
        if header_idx is not None:
            df_body = raw.iloc[header_idx+1:].reset_index()
            df_body.columns = raw.iloc[header_idx]
            df_body = df_body.dropna(how='all').reset_index(drop=True)
            # chuẩn hóa cột
            col_map = {}
            for c in df_body.columns:
                n = normalize_text(c)
                if 'tenhoatchat' in n: col_map[c] = 'Tên hoạt chất'
                elif 'tenthanhphan' in n: col_map[c] = 'Tên hoạt chất'
                elif 'nongdo' in n or 'hamluong' in n: col_map[c] = 'Nồng độ/hàm lượng'
                elif 'nhom' in n and 'thuoc' in n: col_map[c] = 'Nhóm thuốc'
                elif 'soluong' in n: col_map[c] = 'Số lượng'
                elif 'duongdung' in n or 'duong' in n: col_map[c] = 'Đường dùng'
                elif 'gia' in n: col_map[c] = 'Giá kế hoạch'
            df_body.rename(columns=col_map, inplace=True)

            # đánh dấu index gốc để loại trùng
            df_body['_orig_idx'] = df_body['index']
            # chuẩn hóa để merge
            df_body['active_norm'] = df_body['Tên hoạt chất'].apply(normalize_active)
            df_body['conc_norm'] = df_body['Nồng độ/hàm lượng'].apply(normalize_concentration)
            df_body['group_norm'] = df_body['Nhóm thuốc'].apply(normalize_group)
            df2 = file2.copy()
            df2['active_norm'] = df2['Tên hoạt chất'].apply(normalize_active)
            df2['conc_norm'] = df2['Tên hàm lượng'].apply(normalize_concentration)
            df2['group_norm'] = df2['Nhóm thuốc'].apply(normalize_group)
            # merge chính
            merged = pd.merge(df_body, df2, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
            # loại trùng giữ 1 dòng
            merged = merged.drop_duplicates(subset=['_orig_idx'], keep='first')
            # merge file3 temp để bổ sung Địa bàn, Khách hàng
            hosp = file3[file3['Bệnh viện/SYT']==df3_temp['Bệnh viện/SYT'].iloc[0]][['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
            merged = pd.merge(merged, hosp, on='Tên sản phẩm', how='left')

            # chuẩn bị xuất và hiển thị
            export_df = merged.drop(columns=['active_norm','conc_norm','group_norm','_merge','index','_orig_idx'])
            display_df = merged[merged['_merge']=='both'].drop(columns=['active_norm','conc_norm','group_norm','_merge','index','_orig_idx'])
            st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
            st.dataframe(display_df)
            st.session_state['filtered_export'] = export_df.copy()
            st.session_state['filtered_display'] = display_df.copy()

            # tra cứu
            kw = st.text_input("🔍 Tra cứu hoạt chất:")
            if kw:
                df_search = display_df[display_df['Tên hoạt chất'].str.contains(kw, case=False, na=False)]
                st.dataframe(df_search)
            # download
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
                export_df.to_excel(w, index=False, sheet_name='KetQuaLoc')
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
        # định dạng
        def fmt(x):
            if x>=1e9: return f"{x/1e9:.2f} tỷ"
            if x>=1e6: return f"{x/1e6:.2f} triệu"
            if x>=1e3: return f"{x/1e3:.2f} nghìn"
            return str(int(x))

        # chọn nhóm điều trị
        groups = file4['Nhóm điều trị'].dropna().unique()
        sel_group = st.selectbox("Chọn Nhóm điều trị", ['(Tất cả)'] + list(groups))
        if sel_group != '(Tất cả)':
            acts = file4[file4['Nhóm điều trị']==sel_group]['Tên hoạt chất'].unique()
            df = df[df['Tên hoạt chất'].isin(acts)]

        # tổng Trị giá & Số lượng theo hoạt chất
        val_act = df.groupby('Tên hoạt chất')['Trị giá'].sum().reset_index().sort_values('Trị giá', ascending=False)
        val_act['Trị giá'] = val_act['Trị giá'].apply(fmt)
        qty_act = df.groupby('Tên hoạt chất')['Số lượng'].sum().reset_index().sort_values('Số lượng', ascending=False)
        qty_act['Số lượng'] = qty_act['Số lượng'].apply(fmt)
        st.subheader('Tổng Trị giá theo Hoạt chất')
        st.table(val_act)
        st.subheader('Tổng Số lượng theo Hoạt chất')
        st.table(qty_act)

        # Top 10 theo Đường dùng & nhóm điều trị
        st.subheader('Top 10 Hoạt chất theo Đường dùng & Nhóm điều trị (Số lượng)')
        for route in ['tiêm','uống']:
            sub = df[df['Đường dùng'].str.contains(route, case=False, na=False)]
            top = sub.groupby('Tên hoạt chất')['Số lượng'].sum().nlargest(10).reset_index()
            top['Số lượng'] = top['Số lượng'].apply(fmt)
            st.markdown(f"**{route.capitalize()} - Top 10 theo Số lượng**")
            st.table(top)

        # Phân tích theo Khách hàng
        cust = df.groupby('Tên Khách hàng phụ trách triển khai').agg(
            Số_lượng=('Số lượng','sum'),
            Trị_giá=('Trị giá','sum')
        ).reset_index()
        cust['Tỷ lệ SL'] = cust['Số_lượng'] / cust['Số_lượng'].sum()
        cust['Tỷ lệ SL'] = (cust['Tỷ lệ SL']*100).round(2).astype(str) + '%'
        cust['Số_lượng'] = cust['Số_lượng'].apply(fmt)
        cust['Trị_giá'] = cust['Trị_giá'].apply(fmt)
        st.subheader('Phân tích theo Khách hàng phụ trách')
        st.table(cust)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    pass  # Giữ nguyên logic hiện tại

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered_export' not in st.session_state or 'file3_temp' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc' trước.")
    else:
        df_f = st.session_state['filtered_export'].copy()
        df3t = st.session_state['file3_temp'].copy()
        # loại dòng không triển khai
        df3t = df3t[~df3t['Địa bàn'].str.contains('Tạm ngưng triển khai|ko có địa bàn', case=False, na=False)]
        # SL đã trúng
        df_qty = df_f.groupby('Tên sản phẩm')['Số lượng'].sum().rename('SL_trúng').reset_index()
        df_sug = pd.merge(df3t, df_qty, on='Tên sản phẩm', how='left').fillna({'SL_trúng':0})
        # lấy nhóm điều trị
        df_sug = pd.merge(df_sug, file4[['Tên hoạt chất','Nhóm điều trị']], on='Tên hoạt chất', how='left')
        # đề xuất SL: tăng 50% để đạt >50%
        df_sug['Số lượng đề xuất'] = (df_sug['SL_trúng'] * 1.5).apply(np.ceil).astype(int)
        # lý do kết hợp nhóm điều trị
        df_sug['Lý do'] = df_sug.apply(
            lambda r: f"Nhóm {r['Nhóm điều trị']} thường sử dụng các hoạt chất tương ứng; sản phẩm chúng ta thế hệ mới, hiệu quả tốt hơn.",
            axis=1
        )
        st.subheader('File 3 tạm & Đề xuất')
        st.dataframe(df_sug)
        # download
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            df_sug.to_excel(w, index=False, sheet_name='DeXuat')
        st.download_button('⬇️ Tải File Đề Xuất', data=buf.getvalue(), file_name='DeXuat_Thuoc.xlsx')
