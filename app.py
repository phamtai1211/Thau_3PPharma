import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
import unicodedata
import zipfile
from io import BytesIO
from openpyxl import load_workbook

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
        # Strip problematic style/theme files and dataValidations
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

    # Set header and body without user preview
    header = raw.iloc[header_idx].fillna('').astype(str).tolist()
    df_body = raw.iloc[header_idx+1:].copy()
    df_body.columns = header
    df_body = df_body.dropna(subset=header, how='all')
    df_body['_orig_idx'] = df_body.index
    df_body.reset_index(drop=True, inplace=True)

    # Map columns to standard names
    col_map = {}
    for c in df_body.columns:
        n = normalize_text(c)
        if 'tenhoatchat' in n or 'tenthanhphan' in n:
            col_map[c] = 'Tên hoạt chất'
        elif 'nongdo' in n or 'hamluong' in n:
            col_map[c] = 'Nồng độ/hàm lượng'
        elif 'nhom' in n and 'thuoc' in n:
            col_map[c] = 'Nhóm thuốc'
        elif 'soluong' in n:
            col_map[c] = 'Số lượng'
        elif 'duongdung' in n or 'duong' in n:
            col_map[c] = 'Đường dùng'
        elif 'gia' in n:
            col_map[c] = 'Giá kế hoạch'
    df_body.rename(columns=col_map, inplace=True)

    # Normalize reference file2
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

    # Add normalized merge keys
    for df_ in (df_body, df2):
        df_['active_norm'] = df_['Tên hoạt chất'].apply(normalize_active)
        df_['conc_norm'] = df_['Nồng độ/hàm lượng'].apply(normalize_concentration)
        df_['group_norm'] = df_['Nhóm thuốc'].apply(normalize_group)

    # Merge and deduplicate
    merged = pd.merge(df_body, df2, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
    merged.drop_duplicates(subset=['_orig_idx'], keep='first', inplace=True)
    hosp = df3_temp[['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
    merged = pd.merge(merged, hosp, on='Tên sản phẩm', how='left')

    # Prepare display and export DataFrames
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

# === 1. Lọc Danh Mục Thầu ===
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3_temp = file3.copy()
    for col in ['Miền','Vùng','Tỉnh','Bệnh viện/SYT']:
        opts = ['(Tất cả)'] + sorted(df3_temp[col].dropna().unique())
        sel = st.selectbox(f"Chọn {col}", opts)
        if sel != '(Tất cả)':
            df3_temp = df3_temp[df3_temp[col] == sel]

    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx','xls'])
    if uploaded:
        display_df, export_df = process_uploaded(uploaded, df3_temp)
        st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
        st.write(display_df.fillna('').astype(str))

        # lưu session
        st.session_state['filtered_display'] = display_df.copy()
        st.session_state['filtered_export']  = export_df.copy()
        st.session_state['file3_temp']       = df3_temp.copy()

        # fix NameError & tính cột Trị giá
        df = display_df.copy()
        df['Số lượng']     = pd.to_numeric(df['Số lượng'], errors='coerce').fillna(0)
        df['Giá kế hoạch'] = pd.to_numeric(df.get('Giá kế hoạch', 0), errors='coerce').fillna(0)
        df['Trị giá']      = df['Số lượng'] * df['Giá kế hoạch']

        # (các bảng summary nếu có, ví dụ Tổng Trị giá theo Hoạt chất)
        # …

        # nút download kết quả lọc
        from io import BytesIO
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            st.session_state['filtered_export'].to_excel(
                writer, index=False, sheet_name='DanhMucLoc'
            )
        buf.seek(0)
        st.download_button(
            label='⬇️ Tải file Danh Mục Lọc (.xlsx)',
            data=buf.getvalue(),
            file_name='DanhMucLoc.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    # 2.1. Lấy data đã lọc
    df = st.session_state.get('filtered_display', pd.DataFrame()).copy()
    if df.empty:
        st.warning("Bạn chưa thực hiện lọc danh mục. Vui lòng vào tab 'Lọc Danh Mục Thầu' trước.")
    else:
        # 2.2. Định nghĩa hàm format số
        def fmt(x):
            if x >= 1e9: return f"{x/1e9:.2f} tỷ"
            if x >= 1e6: return f"{x/1e6:.2f} triệu"
            if x >= 1e3: return f"{x/1e3:.2f} nghìn"
            return str(int(x))
        
        # 2.3. Chọn Nhóm điều trị
        groups = file4['Nhóm điều trị'].dropna().unique().tolist()
        sel_g = st.selectbox("Chọn Nhóm điều trị", ['(Tất cả)'] + groups)
        if sel_g != '(Tất cả)':
            acts = file4[file4['Nhóm điều trị'] == sel_g]['Tên hoạt chất']
            df = df[df['Tên hoạt chất'].isin(acts)]

        # 2.4. Tính và hiển thị Tổng Trị giá theo Hoạt chất
        val = (
            df
            .groupby('Tên hoạt chất')['Trị giá']
            .sum()
            .reset_index()
            .sort_values('Trị giá', ascending=False)
        )
        val['Trị giá'] = val['Trị giá'].apply(fmt)
        st.subheader("Tổng Trị giá theo Hoạt chất")
        st.table(val)

        # 2.5. Tính và hiển thị Tỷ trọng số lượng theo Hoạt chất
        qty = (
            df
            .groupby('Tên hoạt chất')['Số lượng']
            .sum()
            .reset_index()
            .sort_values('Số lượng', ascending=False)
        )
        total_qty = qty['Số lượng'].sum()
        qty['Tỷ trọng'] = qty['Số lượng'].apply(lambda x: f"{x/total_qty:.2%}")
        st.subheader("Tỷ trọng số lượng theo Hoạt chất")
        st.table(qty)

        # 2.6. Hiển thị tổng số liệu chính
        st.subheader("Chỉ số tổng quan")
        st.metric("Tổng Trị giá", fmt(df['Trị giá'].sum()))
        st.metric("Tổng Số lượng", int(df['Số lượng'].sum()))
        ```

**Hướng dẫn**  
1. Xóa khối cũ `elif option == "Phân Tích Danh Mục Thầu":` đến dòng trước `elif option == "Đề Xuất":`.  
2. Dán đoạn trên với **indent 4 spaces**.  
3. Chạy lại app, phần phân tích sẽ hiển thị dropdown “Chọn Nhóm điều trị” đúng chỗ và các bảng báo cáo.


# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    st.info("Chức năng đang xây dựng...")

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered_export' not in st.session_state or 'file3_temp' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df_f = st.session_state['filtered_export']
        df3t = st.session_state['file3_temp']
        df3t = df3t[~df3t['Địa bàn'].str.contains('Tạm ngưng triển khai|ko có địa bàn', case=False, na=False)]
        qty = df_f.groupby('Tên sản phẩm')['Số lượng'].sum().rename('SL_trúng').reset_index()
        sug = pd.merge(df3t, qty, on='Tên sản phẩm', how='left').fillna({'SL_trúng':0})
        sug = pd.merge(sug, file4[['Tên hoạt chất','Nhóm điều trị']], on='Tên hoạt chất', how='left')
        sug['Số lượng đề xuất'] = (sug['SL_trúng']*1.5).apply(np.ceil).astype(int)
        sug['Lý do'] = sug.apply(lambda r: f"Nhóm {r['Nhóm điều trị']} thường sử dụng; sản phẩm mới, hiệu quả tốt hơn.", axis=1)
        # display with fallback
        try:
            st.dataframe(sug)
        except ValueError:
            st.table(sug)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            sug.to_excel(w, index=False, sheet_name='Đề Xuất')
        st.download_button('⬇️ Tải Đề Xuất', data=buf.getvalue(), file_name='DeXuat.xlsx')
