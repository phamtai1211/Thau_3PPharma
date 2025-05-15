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

# === File processing function ===
def process_uploaded(uploaded, df3_temp):
    # choose sheet with most columns
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

    # auto-detect header row
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
        header_idx = idx if sc>0 else 0
        st.warning(f"Đề xuất dòng tiêu đề: {header_idx+1}")
    st.subheader("🔎 Xem 10 dòng đầu (dòng 1 = index 0)")
    st.dataframe(raw.head(10))
    sel = st.number_input("Chọn dòng header (1-10):", 1, min(10, raw.shape[0]), value=header_idx+1)
    header_idx = sel - 1

    # set header and body
    # Clean header row: replace NaN with empty string
    header = raw.iloc[header_idx].fillna('').astype(str).tolist()
    df_body = raw.iloc[header_idx+1:].copy()
    df_body.columns = header
    df_body = df_body.dropna(subset=header, how='all')
    df_body['_orig_idx'] = df_body.index
    df_body.reset_index(drop=True, inplace=True)

    # map body columns
    col_map = {}
    for c in df_body.columns:
        n = normalize_text(c)
        if 'tenhoatchat' in n or 'tenthanhphan' in n: col_map[c] = 'Tên hoạt chất'
        elif 'nongdo' in n or 'hamluong' in n: col_map[c] = 'Nồng độ/hàm lượng'
        elif 'nhom' in n and 'thuoc' in n: col_map[c] = 'Nhóm thuốc'
        elif 'soluong' in n: col_map[c] = 'Số lượng'
        elif 'duongdung' in n or 'duong' in n: col_map[c] = 'Đường dùng'
        elif 'gia' in n: col_map[c] = 'Giá kế hoạch'
    df_body.rename(columns=col_map, inplace=True)

    # normalize file2
    df2 = file2.copy()
    col_map2 = {}
    for c in df2.columns:
        n = normalize_text(c)
        if 'tenhoatchat' in n: col_map2[c] = 'Tên hoạt chất'
        elif 'nongdo' in n or 'hamluong' in n: col_map2[c] = 'Nồng độ/hàm lượng'
        elif 'nhom' in n and 'thuoc' in n: col_map2[c] = 'Nhóm thuốc'
        elif 'tensanpham' in n: col_map2[c] = 'Tên sản phẩm'
    df2.rename(columns=col_map2, inplace=True)

    # add normalized fields
    for df_ in (df_body, df2):
        df_['active_norm'] = df_['Tên hoạt chất'].apply(normalize_active)
        df_['conc_norm'] = df_['Nồng độ/hàm lượng'].apply(normalize_concentration)
        df_['group_norm'] = df_['Nhóm thuốc'].apply(normalize_group)

    # merge and drop duplicates
    merged = pd.merge(df_body, df2, on=['active_norm','conc_norm','group_norm'], how='left', indicator=True)
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
        if sel != '(Tất cả)': df3_temp = df3_temp[df3_temp[col]==sel]
    uploaded = st.file_uploader("Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        display_df, export_df = process_uploaded(uploaded, df3_temp)
        st.success(f"✅ Tổng dòng khớp: {len(display_df)}")
        st.dataframe(display_df)
        st.session_state['filtered_display'] = display_df
        st.session_state['filtered_export'] = export_df
        kw = st.text_input("🔍 Tra cứu hoạt chất:")
        if kw:
            st.dataframe(display_df[display_df['Tên hoạt chất'].str.contains(kw, case=False)])
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            export_df.to_excel(w, index=False, sheet_name='Kết quả')
        st.download_button('⬇️ Tải File', data=buf.getvalue(), file_name='Ketqua_loc_all.xlsx')

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if 'filtered_display' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df = st.session_state['filtered_display'].copy()
        df['Số lượng'] = pd.to_numeric(df['Số lượng'], errors='coerce').fillna(0)
        df['Giá kế hoạch'] = pd.to_numeric(df.get('Giá kế hoạch',0), errors='coerce').fillna(0)
        df['Trị giá'] = df['Số lượng']*df['Giá kế hoạch']
        def fmt(x):
            if x>=1e9: return f"{x/1e9:.2f} tỷ"
            if x>=1e6: return f"{x/1e6:.2f} triệu"
            if x>=1e3: return f"{x/1e3:.2f} nghìn"
            return str(int(x))
        groups = file4['Nhóm điều trị'].dropna().unique()
        sel_g = st.selectbox("Chọn Nhóm điều trị", ['(Tất cả)']+list(groups))
        if sel_g!='(Tất cả)':
            acts = file4[file4['Nhóm điều trị']==sel_g]['Tên hoạt chất']
            df = df[df['Tên hoạt chất'].isin(acts)]
        val = df.groupby('Tên hoạt chất')['Trị giá'].sum().reset_index().sort_values('Trị giá',False)
        val['Trị giá']=val['Trị giá'].apply(fmt)
        qty= df.groupby('Tên hoạt chất')['Số lượng'].sum().reset_index().sort_values('Số lượng',False)
        qty['Số lượng']=qty['Số lượng'].apply(fmt)
        st.subheader('Tổng Trị giá theo Hoạt chất')
        st.table(val)
        st.subheader('Tổng Số lượng theo Hoạt chất')
        st.table(qty)
        st.subheader('Top 10 theo Đường dùng')
        for r in ['tiêm','uống']:
            sub = df[df['Đường dùng'].str.contains(r, case=False, na=False)]
            topq = sub.groupby('Tên hoạt chất')['Số lượng'].sum().nlargest(10).reset_index()
            topt = sub.groupby('Tên hoạt chất')['Trị giá'].sum().nlargest(10).reset_index()
            topq['Số lượng']=topq['Số lượng'].apply(fmt)
            topt['Trị giá']=topt['Trị giá'].apply(fmt)
            st.markdown(f"**{r.capitalize()} - Top 10 SL**")
            st.table(topq)
            st.markdown(f"**{r.capitalize()} - Top 10 TG**")
            st.table(topt)
        total_sp=df['Tên sản phẩm'].nunique()
        cust=df.groupby('Tên Khách hàng phụ trách triển khai').agg(
            SL=('Số lượng','sum'),TG=('Trị giá','sum'),SP=('Tên sản phẩm',pd.Series.nunique)
        ).reset_index()
        cust['Tỷ lệ SP']=(cust['SP']/total_sp*100).round(2).astype(str)+'%'
        cust['SL']=cust['SL'].apply(fmt)
        cust['TG']=cust['TG'].apply(fmt)
        st.subheader('Phân tích theo Khách hàng phụ trách')
        st.table(cust)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    st.info("Đang xây dựng...")

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered_export' not in st.session_state or 'file3_temp' not in st.session_state:
        st.info("Vui lòng thực hiện 'Lọc Danh Mục Thầu' trước.")
    else:
        df_f=st.session_state['filtered_export']
        df3t=st.session_state['file3_temp']
        df3t=df3t[~df3t['Địa bàn'].str.contains('Tạm ngưng triển khai|ko có địa bàn',case=False,na=False)]
        qty=df_f.groupby('Tên sản phẩm')['Số lượng'].sum().rename('SL_trúng').reset_index()
        sug=pd.merge(df3t,qty,on='Tên sản phẩm',how='left').fillna({'SL_trúng':0})
        sug=pd.merge(sug,file4[['Tên hoạt chất','Nhóm điều trị']],on='Tên hoạt chất',how='left')
        sug['Số lượng đề xuất']=(sug['SL_trúng']*1.5).apply(np.ceil).astype(int)
        sug['Lý do']=sug.apply(lambda r: f"Nhóm {r['Nhóm điều trị']} ... hiệu quả tốt hơn.",axis=1)
        st.dataframe(sug)
        buf=BytesIO()
        with pd.ExcelWriter(buf,engine='xlsxwriter') as w:
            sug.to_excel(w,index=False,sheet_name='DeXuat')
        st.download_button('⬇️ Tải Đề Xuất',data=buf.getvalue(),file_name='DeXuat.xlsx')
