import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO
import plotly.express as px

# --- HELPER FUNCTIONS ---

def find_header(df, keywords, max_rows=20):
    # forward-fill to handle merged cells
    df_ff = df.ffill(axis=0)
    for i in range(min(max_rows, len(df_ff))):
        row = " ".join(df_ff.iloc[i].astype(str).str.lower().str.strip().tolist())
        if all(k.lower() in row for k in keywords):
            return i
    return None


def normalize_active(name):
    s = re.sub(r"\(.*?\)", "", str(name))
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def normalize_conc(conc):
    s = str(conc).lower().replace(',', '.')
    s = s.replace('dung tích','')
    parts = [p.strip() for p in s.split(',') if p.strip()]
    parts = [p for p in parts if re.search(r"\d", p)]
    if len(parts)>=2 and re.search(r"(mg|mcg|g|%)", parts[0]) and 'ml' in parts[-1] and '/' not in parts[0]:
        return parts[0].replace(' ','') + '/' + parts[-1].replace(' ','')
    return ''.join(p.replace(' ','') for p in parts)


def normalize_group(grp):
    return re.sub(r"\D", "", str(grp)).strip()


# --- LOAD DEFAULT DATA ---
@st.cache_data
def load_defaults():
    urls = {
        'file2': 'https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx',
        'file3': 'https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx',
        'file4': 'https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx'
    }
    file2 = pd.read_excel(BytesIO(requests.get(urls['file2']).content))
    file3 = pd.read_excel(BytesIO(requests.get(urls['file3']).content))
    file4 = pd.read_excel(BytesIO(requests.get(urls['file4']).content))
    return file2, file3, file4

file2, file3, file4 = load_defaults()

# filter file3
file3 = file3[~file3['Địa bàn'].astype(str).str.contains("tạm ngưng triển khai|ko có địa bàn", case=False, na=False)]

# Sidebar
st.sidebar.title('Chức năng')
option = st.sidebar.radio('Chọn chức năng', [
    'Lọc Danh Mục Thầu',
    'Phân Tích Danh Mục Thầu',
    'Phân Tích Danh Mục Trúng Thầu',
    'Đề Xuất Hướng Triển Khai',
    'Tra cứu hoạt chất'
])

# --- 1. FILTER TENDER ---
if option=='Lọc Danh Mục Thầu':
    st.header('📂 Lọc Danh Mục Thầu')
    # chọn bệnh viện
    regions = file3['Miền'].dropna().unique().tolist()
    sel_r = st.selectbox('Miền', sorted(regions))
    df3a = file3[file3['Miền']==sel_r]
    areas = df3a['Vùng'].dropna().unique().tolist()
    sel_a = st.selectbox('Vùng', ['(Tất cả)']+sorted(areas))
    if sel_a!='(Tất cả)': df3a = df3a[df3a['Vùng']==sel_a]
    provs = df3a['Tỉnh'].dropna().unique().tolist()
    sel_p = st.selectbox('Tỉnh', sorted(provs))
    df3a = df3a[df3a['Tỉnh']==sel_p]
    hosp = st.selectbox('BV/SYT', df3a['Bệnh viện/SYT'].dropna().unique().tolist())

    upload = st.file_uploader('File Mời Thầu', type=['xlsx'])
    if upload:
        xls = pd.ExcelFile(upload)
        # pick sheet with most cols
        sheet = max(xls.sheet_names, key=lambda s: xls.parse(s,nrows=1,header=None).shape[1])
        raw = pd.read_excel(upload, sheet_name=sheet, header=None)
        hi = find_header(raw, ['tên hoạt chất','số lượng'], max_rows=20)
        if hi is None:
            st.error('Không tìm thấy header trong 20 dòng đầu.')
        else:
            df = raw.ffill(axis=0)
            df = df.iloc[hi+1:]
            df.columns = raw.iloc[hi]
            df = df.reset_index(drop=True)
            # normalize cols
            df['_act'] = df['Tên hoạt chất'].apply(normalize_active)
            df['_conc'] = df['Nồng độ/hàm lượng'].apply(normalize_conc)
            df['_grp'] = df['Nhóm thuốc'].apply(normalize_group)
            # merge
            comp = file2.copy()
            comp['__act'] = comp['Tên hoạt chất'].apply(normalize_active)
            comp['__conc'] = comp['Nồng độ/Hàm lượng'].apply(normalize_conc)
            comp['__grp'] = comp['Nhóm thuốc'].apply(normalize_group)
            merged = df.merge(comp, left_on=['_act','_conc','_grp'],
                             right_on=['__act','__conc','__grp'], how='left', suffixes=('','_cmp'))
            # map hosp info
            hospmap = file3[['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
            merged = merged.merge(hospmap, on='Tên sản phẩm', how='left')
            # calculate SL total per treatment group
            tm = {normalize_active(a):g for a,g in zip(file4['Hoạt chất'],file4['Nhóm điều trị'])}
            totals = {}
            for _,r in df.iterrows():
                g = tm.get(r['_act'],None)
                q = pd.to_numeric(r.get('Số lượng',0),errors='coerce') or 0
                if g: totals[g]=totals.get(g,0)+q
            # add ratio
            ratios=[]
            for _,r in merged.iterrows():
                g = tm.get(r['_act'],None)
                q = pd.to_numeric(r.get('Số lượng',0),errors='coerce') or 0
                if g and totals.get(g,0)>0:
                    ratios.append(f"{q/totals[g]:.2%}")
                else:
                    ratios.append(None)
            merged['Tỷ trọng nhóm thầu'] = ratios
            # UI show only matched
            df_matched = merged[~merged['Tên sản phẩm_cmp'].isna()]
            st.success(f'✅ Đã lọc {len(df_matched)} dòng phù hợp.')
            st.dataframe(df_matched, height=400)
            # download full
            buf=BytesIO()
            writer=pd.ExcelWriter(buf,engine='xlsxwriter')
            merged.to_excel(writer,index=False,sheet_name='Full')
            writer.save()
            st.download_button('⬇️ Xuất full kết quả',buf.getvalue(), 'ketqua.xlsx')
            st.session_state['filtered']=merged
            st.session_state['matched']=df_matched
            st.session_state['hospital']=hosp

# ... tiếp các chức năng 2,3,4,5 tương tự với UI và placeholder logic ...

else:
    st.info('Chức năng đang được phát triển.')
