import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO
import plotly.express as px

# --- Utility functions ---
def find_header_row(df, keywords):
    for i in range(min(20, len(df))):
        row = " ".join(df.iloc[i].astype(str).tolist()).lower()
        if all(k in row for k in keywords): return i
    raise ValueError(f"Không tìm thấy header trong 20 dòng đầu chứa: {keywords}")

def normalize_active(s):
    s = re.sub(r"\(.*?\)", "", str(s))
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def normalize_conc(s):
    s = str(s).lower().replace(',', '.')
    s = re.sub(r'dung tích', '', s)
    parts = [p.strip() for p in s.split(',') if p.strip()]
    parts = [p for p in parts if re.search(r'\d', p)]
    if len(parts)>=2 and re.search(r'(mg|g|mcg|%)', parts[0]) and 'ml' in parts[-1]:
        return parts[0].replace(' ','') + '/' + parts[-1].replace(' ','')
    return ''.join(p.replace(' ','') for p in parts)

def normalize_group(s):
    return re.sub(r'\D','', str(s))

@st.cache_data
def load_defaults():
    urls = [
        "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx",
    ]
    file2 = pd.read_excel(BytesIO(requests.get(urls[0]).content))
    file3 = pd.read_excel(BytesIO(requests.get(urls[1]).content))
    file4 = pd.read_excel(BytesIO(requests.get(urls[2]).content))
    return file2, file3, file4

file2, file3, file4 = load_defaults()

# Temporary filtered file3 for proposals
def get_file3_temp():
    df = file3.copy()
    df['Địa bàn'] = df['Địa bàn'].fillna('').astype(str)
    return df[~df['Địa bàn'].str.contains('tạm ngưng triển khai|ko có địa bàn', case=False)]
file3_temp = get_file3_temp()

# Sidebar
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai"
])

# 1. Lọc Danh Mục Thầu
if option=="Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    # select filters on file3
    regi = st.selectbox("Chọn Miền", sorted(file3['Miền'].dropna().unique()))
    df3 = file3[file3['Miền']==regi]
    areas = sorted(df3['Vùng'].dropna().unique())
    area = st.selectbox("Chọn Vùng", ['(Tất cả)']+areas)
    if area!='(Tất cả)': df3=df3[df3['Vùng']==area]
    provs = sorted(df3['Tỉnh'].dropna().unique())
    prov = st.selectbox("Chọn Tỉnh", provs)
    df3=df3[df3['Tỉnh']==prov]
    hosp = st.selectbox("Chọn BV/SYT", sorted(df3['Bệnh viện/SYT'].dropna().unique()))
    uploaded = st.file_uploader("File mời thầu (.xlsx)", type=['xlsx'])
    if uploaded:
        xls = pd.ExcelFile(uploaded)
        sheet = max(xls.sheet_names, key=lambda s: xls.parse(s,nrows=1,header=None).shape[1])
        raw = pd.read_excel(uploaded, sheet_name=sheet, header=None)
        try:
            hdr = find_header_row(raw, ['tên hoạt chất','số lượng'])
        except Exception as e:
            st.error(str(e))
            st.stop()
        df = raw.iloc[hdr+1:].copy().reset_index(drop=True)
        df.columns = raw.iloc[hdr].tolist()
        df = df.dropna(how='all').reset_index(drop=True)
        # detect columns
        act = next(c for c in df.columns if 'hoạt chất' in c.lower())
        conc = next(c for c in df.columns if 'hàm lượng' in c.lower() or 'nồng độ' in c.lower())
        grp = next(c for c in df.columns if 'nhóm' in c.lower())
        # normalize
        df['_act']=df[act].apply(normalize_active)
        df['_conc']=df[conc].apply(normalize_conc)
        df['_grp']=df[grp].apply(normalize_group)
        comp=file2.copy()
        comp['_act']=comp['Tên hoạt chất'].apply(normalize_active)
        comp['_conc']=comp['Nồng độ/Hàm lượng'].apply(normalize_conc)
        comp['_grp']=comp['Nhóm thuốc'].apply(normalize_group)
        merged=pd.merge(df, comp, on=['_act','_conc','_grp'], how='left', suffixes=('','_cmp'))
        # attach hosp info
        hosp_data=file3[file3['Bệnh viện/SYT']==hosp][['Tên sản phẩm','Địa bàn','Tên Khách hàng phụ trách triển khai']]
        merged=pd.merge(merged, hosp_data, on='Tên sản phẩm', how='left')
        # drop duplicates
        out=merged.drop_duplicates(['_act','_conc','_grp'])
        # compute ratio
        tmap={normalize_active(a):g for a,g in zip(file4['Hoạt chất'], file4['Nhóm điều trị'])}
        df['qty']=pd.to_numeric(df.get('Số lượng',0),errors='coerce').fillna(0)
        totals=df.groupby(df['_act']).qty.sum()
        out['SL']=pd.to_numeric(out.get('Số lượng',0),errors='coerce').fillna(0)
        out['Tỷ trọng nhóm thầu']=out['_act'].map(lambda a: totals.get(a,0))
        out['Tỷ trọng nhóm thầu']= (out['SL']/out['Tỷ trọng nhóm thầu']).fillna(0).map(lambda x:f"{x:.2%}")
        st.success(f"✅ Đã lọc xong {len(out)} dòng.")
        st.dataframe(out)
        buf=BytesIO(); out.to_excel(buf,index=False)
        st.download_button("⬇️ Download kết quả", data=buf.getvalue(), file_name='ketqua_loc.xlsx')
        st.session_state['filtered']=out; st.session_state['hosp']=hosp

# 2. Phân Tích Danh Mục Thầu
elif option=="Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if 'filtered' not in st.session_state:
        st.info("Làm bước Lọc Danh Mục Thầu trước.")
    else:
        df=st.session_state['filtered'].copy()
        df['Số lượng']=pd.to_numeric(df['Số lượng'],errors='coerce').fillna(0)
        df['Giá kế hoạch']=pd.to_numeric(df.get('Giá kế hoạch',0),errors='coerce').fillna(0)
        df['Trị giá']=df['Số lượng']*df['Giá kế hoạch']
        # four charts: Top10 injection/value & ingestion/value
        df['route']=df['Đường dùng'].str.lower().apply(lambda x:'Tiêm' if 'tiêm' in x else ('Uống' if 'uống' in x else 'Khác'))
        cols={'Tiêm':'Tiêm','Uống':'Uống'}
        for rt in ['Tiêm','Uống']:
            for metric in ['Số lượng','Trị giá']:
                sub=df[df['route']==rt]
                top=sub.groupby('Tên hoạt chất')[metric].sum().nlargest(10).reset_index()
                fig=px.bar(top, x='Tên hoạt chất', y=metric, title=f"Top10 {rt} theo {metric}")
                st.plotly_chart(fig, use_container_width=True)
        # group treatment
        tmap={normalize_active(a):g for a,g in zip(file4['Hoạt chất'], file4['Nhóm điều trị'])}
        df['Nhóm điều trị']=df['_act'].map(tmap).fillna('Khác')
        # by value
        tv=df.groupby('Nhóm điều trị')['Trị giá'].sum().reset_index().sort_values('Trị giá',ascending=False)
        st.plotly_chart(px.bar(tv, x='Nhóm điều trị', y='Trị giá', title='Trị giá theo nhóm điều trị'), use_container_width=True)
        # choose group for top10 active by value
        grp=st.selectbox("Chọn nhóm điều trị", tv['Nhóm điều trị'])
        sub=df[df['Nhóm điều trị']==grp]
        topv=sub.groupby('Tên hoạt chất')['Trị giá'].sum().nlargest(10).reset_index()
        st.plotly_chart(px.bar(topv, x='Tên hoạt chất', y='Trị giá', title=f"Top10 HC trong nhóm {grp}"), use_container_width=True)
        # by quantity
        tq=df.groupby('Nhóm điều trị')['Số lượng'].sum().reset_index().sort_values('Số lượng',ascending=False)
        st.plotly_chart(px.bar(tq, x='Nhóm điều trị', y='Số lượng', title='SL theo nhóm điều trị'), use_container_width=True)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option=="Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win=st.file_uploader("File kết quả trúng (.xlsx)",type=['xlsx'])
    inv=st.file_uploader("File mời thầu để đối chiếu (tùy chọn)",type=['xlsx'])
    if win:
        xls=pd.ExcelFile(win)
        sheet=max(xls.sheet_names, key=lambda s:xls.parse(s,nrows=1,header=None).shape[1])
        raw=pd.read_excel(win, sheet_name=sheet, header=None)
        try:
            h=find_header_row(raw,['tên hoạt chất','nhà thầu trúng'])
        except:
            st.error("Không xác định header trúng thầu."); st.stop()
        dfw=raw.iloc[h+1:].reset_index(drop=True)
        dfw.columns=raw.iloc[h].tolist(); dfw=dfw.dropna(how='all')
        dfw['Số lượng']=pd.to_numeric(dfw.get('Số lượng',0),errors='coerce').fillna(0)
        price_col=next((c for c in dfw.columns if 'Giá trúng' in str(c)), 'Giá kế hoạch')
        dfw[price_col]=pd.to_numeric(dfw.get(price_col,0),errors='coerce').fillna(0)
        dfw['Trị giá']=dfw['Số lượng']*dfw[price_col]
        top= dfw.groupby('Nhà thầu trúng')['Trị giá'].sum().nlargest(20).reset_index()
        st.plotly_chart(px.bar(top, x='Nhà thầu trúng', y='Trị giá', orientation='h', title='Top20 nhà thầu'), use_container_width=True)
        # treatment pie
        tmap={normalize_active(a):g for a,g in zip(file4['Hoạt chất'], file4['Nhóm điều trị'])}
        dfw['_act']=dfw['Tên hoạt chất'].apply(normalize_active)
        dfw['Nhóm điều trị']=dfw['_act'].map(tmap).fillna('Khác')
        tw=dfw.groupby('Nhóm điều trị')['Trị giá'].sum().reset_index()
        st.plotly_chart(px.pie(tw, names='Nhóm điều trị', values='Trị giá', title='Cơ cấu trị giá trúng'), use_container_width=True)

# 4. Đề Xuất Hướng Triển Khai
elif option=="Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if 'filtered' not in st.session_state:
        st.info("Hãy chạy Lọc Danh Mục Thầu trước.")
    else:
        df=st.session_state['filtered']
        hosp=st.session_state['hosp']
        # SL đã làm: từ file3_temp
        df3=file3_temp[file3_temp['Bệnh viện/SYT']==hosp]
        df3['_act']=df3['Hoạt chất'].apply(normalize_active)
        df3['_conc']=df3['Hàm lượng'].apply(normalize_conc)
        df3['_grp']=df3['Nhóm thuốc'].apply(normalize_group)
        done=df3.groupby(['_act','_conc','_grp'])['SL thực tế'].sum()
        # SL thầu: từ df
        planned=df.set_index(['_act','_conc','_grp'])['Số lượng']
        idx=planned.index.union(done.index)
        prop=pd.DataFrame(index=idx)
        prop['Đã làm']=done; prop['Thầu yêu cầu']=planned
        prop=prop.fillna(0)
        prop['Đề xuất năm sau']=(prop['Thầu yêu cầu']-prop['Đã làm']).clip(lower=0).astype(int)
        prop=prop.reset_index()
        st.dataframe(prop)
        buf=BytesIO(); prop.to_excel(buf,index=False)
        st.download_button("⬇️ Download đề xuất", data=buf.getvalue(), file_name='de_xuat.xlsx')
