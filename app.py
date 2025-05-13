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
    url_file2 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx"
    url_file3 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx"
    url_file4 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"

    file2 = pd.read_excel(BytesIO(requests.get(url_file2).content))
    file3 = pd.read_excel(BytesIO(requests.get(url_file3).content))
    file4 = pd.read_excel(BytesIO(requests.get(url_file4).content))
    return file2, file3, file4

file2, file3, file4 = load_default_data()

# Chuẩn hóa xóa dấu
def remove_diacritics(s: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

# Chuẩn hóa chuỗi để so sánh
def normalize_text(s: str) -> str:
    s = str(s)
    s = remove_diacritics(s).lower()
    s = re.sub(r'\s+', '', s)
    return s

# Hàm chuẩn hóa hoạt chất, hàm lượng, nhóm

def normalize_active(name: str) -> str:
    return re.sub(r'\s+', ' ', re.sub(r'\(.*?\)', '', str(name))).strip().lower()

def normalize_concentration(conc: str) -> str:
    s = str(conc).lower().replace(',', '.')
    parts = [p.strip() for p in re.split(r'[;,]', s) if p.strip()]
    parts = [p for p in parts if re.search(r'\d', p)]
    if len(parts) >= 2 and re.search(r'(mg|mcg|g|%)', parts[0]) and 'ml' in parts[-1] and '/' not in parts[0]:
        return parts[0].replace(' ', '') + '/' + parts[-1].replace(' ', '')
    return ''.join([p.replace(' ', '') for p in parts])

def normalize_group(grp: str) -> str:
    return re.sub(r'\D', '', str(grp)).strip()

# Sidebar: chức năng chính
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", 
    ["Lọc Danh Mục Thầu", "Phân Tích Danh Mục Thầu", "Phân Tích Danh Mục Trúng Thầu", "Đề Xuất Hướng Triển Khai"] )

# 1. Lọc Danh Mục Thầu
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    # Tiết lưu vùng chọn để dùng later
    regions = sorted(file3["Miền"].dropna().unique())
    selected_region = st.selectbox("Chọn Miền", ["(Tất cả)"] + regions)
    df3_sel = file3 if selected_region == "(Tất cả)" else file3[file3["Miền"] == selected_region]
    areas = sorted(df3_sel["Vùng"].dropna().unique())
    selected_area = st.selectbox("Chọn Vùng", ["(Tất cả)"] + areas) if areas else None
    if selected_area and selected_area != "(Tất cả)": df3_sel = df3_sel[df3_sel["Vùng"] == selected_area]
    provinces = sorted(df3_sel["Tỉnh"].dropna().unique())
    selected_prov = st.selectbox("Chọn Tỉnh", ["(Tất cả)"] + provinces)
    if selected_prov and selected_prov != "(Tất cả)": df3_sel = df3_sel[df3_sel["Tỉnh"] == selected_prov]
    hospitals = sorted(df3_sel["Bệnh viện/SYT"].dropna().unique())
    selected_hospital = st.selectbox("Chọn Bệnh viện/Sở Y Tế", ["(Tất cả)"] + hospitals)
    if selected_hospital and selected_hospital != "(Tất cả)": df3_sel = df3_sel[df3_sel["Bệnh viện/SYT"] == selected_hospital]
    # Lưu file3 temp
    st.session_state["file3_temp"] = df3_sel.copy()

    uploaded_file = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=["xlsx"])
    if uploaded_file and (selected_hospital and selected_hospital != "(Tất cả)"):
        xls = pd.ExcelFile(uploaded_file)
        # chọn sheet có nhiều cột nhất
        sheet_name = max(xls.sheet_names, key=lambda name: pd.read_excel(uploaded_file, sheet_name=name, nrows=1, header=None).shape[1])
        df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        # tìm header row (1–10) bỏ qua merge
        header_idx = None
        score = []
        for i in range(min(10, df_raw.shape[0])):
            row = df_raw.iloc[i].fillna('').astype(str).tolist()
            text = ' '.join(row)
            norm = normalize_text(text)
            # kiểm tra đủ các trường
            if 'tenhoatchat' in norm and ('soluong' in norm or 'nobanthe' in norm):
                header_idx = i; break
            # tính điểm
            sc = ('tenhoatchat' in norm) + ('soluong' in norm) + ('nhomthuoc' in norm) + ('nongdohamluong' in norm)
            score.append((i, sc))
        if header_idx is None:
            # lấy row có điểm cao nhất nếu có
            idx, sc = max(score, key=lambda x: x[1])
            if sc > 0:
                header_idx = idx
                st.warning(f"Tự động chọn dòng tiêu đề tại dòng {idx+1}")
            else:
                st.error("❌ Không xác định được dòng tiêu đề trong file.")
        if header_idx is not None:
            header = df_raw.iloc[header_idx].tolist()
            df_all = df_raw.iloc[header_idx+1:].reset_index(drop=True)
            df_all.columns = header
            df_all = df_all.dropna(how='all').reset_index(drop=True)
            # chuẩn hóa tên cột
            col_map = {}
            for col in df_all.columns:
                n = normalize_text(col)
                if 'tenhoatchat' in n:
                    col_map[col] = 'Tên hoạt chất'
                elif 'tênhoatchat' in n and 'tenthanhphan' in n:
                    col_map[col] = 'Tên hoạt chất'
                elif 'nongdo' in n or 'hamluong' in n:
                    col_map[col] = 'Nồng độ/hàm lượng'
                elif 'nhom' in n and 'thuoc' in n:
                    col_map[col] = 'Nhóm thuốc'
                elif 'soluong' in n:
                    col_map[col] = 'Số lượng'
                elif 'gia' in n:
                    col_map[col] = 'Giá kế hoạch'
                elif 'duongdung' in n or 'duong' in n:
                    col_map[col] = 'Đường dùng'
            df_all.rename(columns=col_map, inplace=True)

            # chuẩn bị so sánh
            df_all['active_norm'] = df_all['Tên hoạt chất'].apply(normalize_active)
            df_all['conc_norm'] = df_all['Nồng độ/hàm lượng'].apply(normalize_concentration)
            df_all['group_norm'] = df_all['Nhóm thuốc'].apply(normalize_group)
            df_comp = file2.copy()
            df_comp['active_norm'] = df_comp['Tên hoạt chất'].apply(normalize_active)
            df_comp['conc_norm'] = df_comp['Nồng độ/Hàm lượng'].apply(normalize_concentration)
            df_comp['group_norm'] = df_comp['Nhóm thuốc'].apply(normalize_group)
            # merge left để giữ mọi dòng
            merged = pd.merge(df_all, df_comp, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True, suffixes=('','_comp'))
            # merge với file3 temp để lấy địa bàn, khách hàng
            hosp_data = file3[file3['Bệnh viện/SYT']==selected_hospital][['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
            merged = pd.merge(merged, hosp_data, on='Tên sản phẩm', how='left')
            # lưu xuất
            export_df = merged.drop(columns=['active_norm','conc_norm','group_norm','_merge'])
            display_df = merged[merged['_merge']=='both'].drop(columns=['active_norm','conc_norm','group_norm','_merge'])
            st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
            st.dataframe(display_df)
            # lưu
            st.session_state['filtered_export'] = export_df.copy()
            st.session_state['filtered_display'] = display_df.copy()

            # cho tra cứu hoạt chất
            kw = st.text_input("🔍 Tra cứu hoạt chất (nhập tên) để lọc kết quả:")
            if kw:
                kw_norm = kw.strip().lower()
                df_search = display_df[display_df['Tên hoạt chất'].str.lower().str.contains(kw_norm)]
                st.dataframe(df_search)

            # tải về file
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                export_df.to_excel(writer, index=False, sheet_name='KetQuaLoc')
            st.download_button('⬇️ Tải File Kết Quả', data=buf.getvalue(), file_name='Ketqua_loc_all.xlsx')
            # lưu session
            st.session_state['filtered_df'] = display_df.copy()

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu (Số liệu)")
    if 'filtered_df' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df = st.session_state['filtered_df'].copy()
        df['Số lượng'] = pd.to_numeric(df['Số lượng'], errors='coerce').fillna(0)
        df['Giá kế hoạch'] = pd.to_numeric(df.get('Giá kế hoạch',0), errors='coerce').fillna(0)
        df['Trị giá'] = df['Số lượng'] * df['Giá kế hoạch']
        # Hàm định dạng số
        def fmt(x):
            if x>=1e9: return f"{x/1e9:.2f} tỷ"
            if x>=1e6: return f"{x/1e6:.2f} triệu"
            if x>=1e3: return f"{x/1e3:.2f} nghìn"
            return str(x)
        # Tổng trị giá theo hoạt chất
        val_act = df.groupby('Tên hoạt chất')['Trị giá'].sum().reset_index().sort_values('Trị giá',ascending=False)
        val_act['Trị giá'] = val_act['Trị giá'].apply(fmt)
        st.subheader('Tổng Trị giá theo Hoạt chất')
        st.table(val_act)
        # Tổng số lượng theo hoạt chất
        qty_act = df.groupby('Tên hoạt chất')['Số lượng'].sum().reset_index().sort_values('Số lượng',ascending=False)
        qty_act['Số lượng'] = qty_act['Số lượng'].apply(fmt)
        st.subheader('Tổng Số lượng theo Hoạt chất')
        st.table(qty_act)
        # Phân tích theo đường dùng (Tiêm & Uống)
        routes = {'Tiêm':'tiêm','Uống':'uống'}
        st.subheader('Top 10 Hoạt chất theo từng Đường dùng')
        for label, key in routes.items():
            sub = df[df['Đường dùng'].str.contains(key, case=False, na=False)]
            top_qty = sub.groupby('Tên hoạt chất')['Số lượng'].sum().nlargest(10).reset_index()
            top_val = sub.groupby('Tên hoạt chất')['Trị giá'].sum().nlargest(10).reset_index()
            top_val['Trị giá'] = top_val['Trị giá'].apply(fmt)
            st.markdown(f"**{label} - Top 10 theo Số lượng**")
            st.table(top_qty)
            st.markdown(f"**{label} - Top 10 theo Trị giá**")
            st.table(top_val)
        # Phân tích theo khách hàng phụ trách
        rep = df.groupby('Tên Khách hàng phụ trách triển khai').agg({'Số lượng':'sum','Trị giá':'sum'}).reset_index().sort_values('Trị giá',ascending=False)
        rep['Số lượng'] = rep['Số lượng'].apply(fmt)
        rep['Trị giá'] = rep['Trị giá'].apply(fmt)
        st.subheader('Phân tích theo Khách hàng phụ trách')
        st.table(rep)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    # (Giữ nguyên logic hiện có)
    pass

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered_df' not in st.session_state or 'file3_temp' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc' và 'Phân tích' trước.")
    else:
        df_f = st.session_state['filtered_export']
        df3_temp = st.session_state['file3_temp'].copy()
        # loại các dòng không triển khai
        df3_temp = df3_temp[~df3_temp['Địa bàn'].str.contains('Tạm ngưng triển khai|ko có địa bàn', case=False, na=False)]
        # tính số lượng đã trúng
        df_qty = df_f.groupby('Tên sản phẩm')['Số lượng'].sum().rename('SL_trung').reset_index()
        df_sug = pd.merge(df3_temp, df_qty, on='Tên sản phẩm', how='left').fillna({'SL_trung':0})
        # đề xuất số lượng: tăng 50% so với SL_trung để đạt tỷ trọng >50%
        df_sug['Số lượng đề xuất'] = (df_sug['SL_trung'] * 1.5).apply(np.ceil).astype(int)
        df_sug['Lý do'] = 'Tăng 50% so với lần trước để đạt tỷ trọng >50%'
        st.subheader('File 3 tạm đã lọc & đề xuất')
        st.dataframe(df_sug)
        # tải về
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            df_sug.to_excel(w, index=False, sheet_name='DeXuat')
        st.download_button('⬇️ Tải File Đề Xuất', data=buf.getvalue(), file_name='DeXuat_Thuoc.xlsx')
