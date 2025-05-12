import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import requests
from io import BytesIO
import plotly.express as px

# ————— Helper —————
def remove_accents(s):
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))

def norm(s):
    s = remove_accents(str(s)).lower().strip()
    return re.sub(r"\s+", " ", s)

def find_header(df, keys):
    for i in range(10):
        row = " ".join(df.iloc[i].fillna("").astype(str)).lower()
        if all(k in row for k in keys):
            return i
    return None

@st.cache_data
def load_data():
    urls = {
        "file2":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        "file3":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        "file4":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    }
    f2 = pd.read_excel(BytesIO(requests.get(urls["file2"]).content))
    f3 = pd.read_excel(BytesIO(requests.get(urls["file3"]).content))
    f4 = pd.read_excel(BytesIO(requests.get(urls["file4"]).content))
    return f2, f3, f4

file2, file3, file4 = load_data()
# tìm cột hàm lượng
conc_col = next((c for c in file3.columns if "nồng độ" in c.lower()), None)

# ——— lọc file3 ———
file3["Địa bàn"] = file3["Địa bàn"].fillna("")
file3 = file3[~file3["Địa bàn"].str.contains("tạm ngưng triển khai|ko có địa bàn", case=False)]
st.session_state["file3"] = file3

# ——— Sidebar ———
st.sidebar.title("Chức năng")
opt = st.sidebar.radio("", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai"
])

# ——— 1. LỌC DANH MỤC ———
if opt=="Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3 = st.session_state["file3"].copy()
    # chọn Miền/Vùng/Tỉnh/BV
    r = st.selectbox("Miền", ["(Tất cả)"]+sorted(df3["Miền"].dropna().unique()))
    if r!="(Tất cả)": df3=df3[df3["Miền"]==r]
    a = st.selectbox("Vùng", ["(Tất cả)"]+sorted(df3["Vùng"].dropna().unique()))
    if a!="(Tất cả)": df3=df3[df3["Vùng"]==a]
    p = st.selectbox("Tỉnh", ["(Tất cả)"]+sorted(df3["Tỉnh"].dropna().unique()))
    if p!="(Tất cả)": df3=df3[df3["Tỉnh"]==p]
    h = st.selectbox("BV/SYT", sorted(df3["Bệnh viện/SYT"].unique()))

    up = st.file_uploader("File mời thầu", type="xlsx")
    if not up:
        st.info("Tải lên file trước đã.")
        st.stop()

    xls = pd.ExcelFile(up)
    sheet = max(xls.sheet_names, key=lambda n: xls.parse(n,nrows=1,header=None).shape[1])
    raw = pd.read_excel(up, sheet_name=sheet, header=None)
    hdr = find_header(raw, ["tên hoạt chất","số lượng"])
    if hdr is None:
        st.error("Không tìm được header.")
        st.stop()
    cols = raw.iloc[hdr].tolist()
    df = raw.iloc[hdr+1:].copy().reset_index(drop=True)
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)

    # chuẩn hóa key
    df["_act"] = df["Tên hoạt chất"].apply(norm)
    df["_conc"]= df[conc_col].apply(norm)
    df["_grp"]= df["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))

    c2 = file2.copy()
    c2["_act"]= c2["Tên hoạt chất"].apply(norm)
    c2["_conc"]= c2["Nồng độ/Hàm lượng"].apply(norm)
    c2["_grp"]= c2["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))

    # merge left giữ nguyên dòng gốc
    m = pd.merge(df, c2, on=["_act","_conc","_grp"], how="left", suffixes=("","_cmp"))

    # gắn in4 BV/hàng hóa
    info = df3[df3["Bệnh viện/SYT"]==h][["Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai"]].drop_duplicates()
    m = pd.merge(m, info, on="Tên sản phẩm", how="left")

    # tính tỷ trọng nhóm thầu (chỉ 5 nhóm thuốc có mặt)
    total_grp = m.groupby("Nhóm thuốc")["Số lượng"].transform("sum")
    m["Tỷ trọng nhóm thầu"] = (pd.to_numeric(m["Số lượng"],errors="coerce").fillna(0)/total_grp).fillna(0).apply(lambda x:f"{x:.2%}")

    st.session_state["filtered"] = m
    st.session_state["df_all"] = df

    st.success(f"✅ Đã lọc xong {len(m)} dòng.")
    # --- Hiển thị không trùng: drop duplicates theo key (_act, _conc, _grp) ---
    display = m.drop_duplicates(subset=["_act","_conc","_grp"]).copy()
    # Chỉ show những dòng có sản phẩm (Tên sản phẩm không null)
    display = display[display["Tên sản phẩm"].notna()]
    st.dataframe(display, height=500)

    buf = BytesIO()
    m.to_excel(buf, index=False, sheet_name="KetQuaLoc")
    st.download_button("⬇️ Tải về full kết quả", buf.getvalue(), "KetQuaLoc.xlsx")

# ——— 2. PHÂN TÍCH DANH MỤC ———
elif opt=="Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if "df_all" not in st.session_state:
        st.info("Chạy Lọc trước.")
        st.stop()
    dfA = st.session_state["df_all"].copy()
    dfA["Số lượng"]=pd.to_numeric(dfA["Số lượng"],errors="coerce").fillna(0)
    dfA["Giá kế hoạch"]=pd.to_numeric(dfA.get("Giá kế hoạch",0),errors="coerce").fillna(0)
    dfA["Trị giá"]=dfA["Số lượng"]*dfA["Giá kế hoạch"]

    # tích hợp Tra cứu
    term = st.text_input("Tra cứu hoạt chất")
    if term:
        dfA = dfA[dfA["Tên hoạt chất"].str.contains(term, case=False, na=False)]

    # format number
    pd.options.display.float_format = '{:,.0f}'.format

    def plot(df, x,y,title):
        fig=px.bar(df, x=x,y=y,title=title)
        fig.update_traces(texttemplate="%{y:,.0f}",textposition="outside")
        st.plotly_chart(fig,use_container_width=True)

    # Trị giá theo nhóm thuốc
    gv = dfA.groupby("Nhóm thuốc")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False)
    plot(gv,"Nhóm thuốc","Trị giá","Trị giá theo Nhóm thuốc")

    # Trị giá theo Đường dùng
    dfA["Đường"]=dfA["Đường dùng"].apply(lambda s:"Tiêm" if "tiêm" in str(s).lower() else("Uống" if "uống" in str(s).lower() else "Khác"))
    g2=dfA.groupby("Đường")["Trị giá"].sum().reset_index()
    plot(g2,"Đường","Trị giá","Trị giá theo Đường dùng")

    # Top 10 HC theo SL & Trị giá
    top_sl = dfA.groupby("Tên hoạt chất")["Số lượng"].sum().reset_index().sort_values("Số lượng",ascending=False).head(10)
    top_v  = dfA.groupby("Tên hoạt chất")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False).head(10)
    plot(top_sl,"Tên hoạt chất","Số lượng","Top 10 HC theo SL")
    plot(top_v, "Tên hoạt chất","Trị giá","Top 10 HC theo Trị giá")

    # Phân tích Nhóm điều trị theo số lượng & trị giá
    treat = {norm(a):g for a,g in zip(file4["Hoạt chất"],file4["Nhóm điều trị"])}
    dfA["Nhóm điều trị"]=dfA["Tên hoạt chất"].apply(lambda x:treat.get(norm(x),"Khác"))
    t2 = dfA.groupby("Nhóm điều trị")["Số lượng"].sum().reset_index().sort_values("Số lượng",ascending=False)
    plot(t2,"Nhóm điều trị","Số lượng","SL mời thầu theo Nhóm điều trị")
    sel = st.selectbox("Chọn nhóm điều trị xem Top 10 HC (Trị giá)", t2["Nhóm điều trị"].tolist())
    if sel:
        t3 = dfA[dfA["Nhóm điều trị"]==sel].groupby("Tên hoạt chất")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False).head(10)
        plot(t3,"Tên hoạt chất","Trị giá",f"Top 10 HC trị giá - {sel}")

# ——— 3. PHÂN TÍCH TRÚNG THẦU ———
elif opt=="Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win=st.file_uploader("File Kết quả Trúng Thầu",type="xlsx")
    if not win: st.info("Tải file lên"); st.stop()
    xlsw=pd.ExcelFile(win)
    sh=max(xlsw.sheet_names,key=lambda n:xlsw.parse(n,nrows=1,header=None).shape[1])
    rw=pd.read_excel(win,sheet_name=sh,header=None)
    hi=find_header(rw,["tên hoạt chất","nhà thầu trúng"])
    if hi is None: st.error("No header"); st.stop()
    hdr=rw.iloc[hi].tolist(); dfw=rw.iloc[hi+1:].reset_index(drop=True); dfw.columns=hdr; dfw=dfw.dropna(how="all")
    dfw["Số lượng"]=pd.to_numeric(dfw.get("Số lượng",0),errors="coerce").fillna(0)
    pc=next((c for c in dfw.columns if "Giá trúng" in c),"Giá kế hoạch")
    dfw[pc]=pd.to_numeric(dfw.get(pc,0),errors="coerce").fillna(0)
    dfw["Trị giá"]=dfw["Số lượng"]*dfw[pc]
    wv=dfw.groupby("Nhà thầu trúng")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False).head(20)
    f1=px.bar(wv,x="Trị giá",y="Nhà thầu trúng",orientation="h",title="Top 20 Nhà thầu"); f1.update_traces(texttemplate="%{x:,.0f}",textposition="outside"); st.plotly_chart(f1)
    dfw["Nhóm điều trị"]=dfw["Tên hoạt chất"].apply(lambda x:treat.get(norm(x),"Khác"))
    tw=dfw.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index()
    f2=px.pie(tw,names="Nhóm điều trị",values="Trị giá",title="Cơ cấu trúng thầu"); st.plotly_chart(f2)

# ——— 4. ĐỀ XUẤT HƯỚNG TRIỂN KHAI ———
elif opt=="Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if "filtered" not in st.session_state:
        st.info("Chạy Lọc trước."); st.stop()
    dfm=st.session_state["filtered"]
    # SL đã thực hiện
    done=dfm.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Đã làm"})
    # SL yêu cầu BV
    req=file3_filtered.copy()
    req["_act"]=req["Tên hoạt chất"].apply(norm)
    req["_conc"]=req[conc_col].apply(norm)
    req["_grp"]=req["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))
    req=req.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_YeuCau"})
    # merge & tính đề xuất
    sug=pd.merge(req,done,on=["_act","_conc","_grp"],how="left").fillna(0)
    sug["Đề xuất"]= (sug["SL_YeuCau"]-sug["SL_Đã làm"]).clip(lower=0).astype(int)
    # thêm khách hàng phụ trách
    kh=file3_filtered.copy()
    kh["_act"]=kh["Tên hoạt chất"].apply(norm)
    kh["_conc"]=kh[conc_col].apply(norm)
    kh["_grp"]=kh["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))
    kh=kh.groupby(["_act","_conc","_grp"])["Tên Khách hàng phụ trách triển khai"].first().reset_index()
    sug=pd.merge(sug,kh,on=["_act","_conc","_grp"],how="left")
    st.dataframe(sug, height=500)
