import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import requests
from io import BytesIO
import plotly.express as px

# ——— Helper functions ———
def remove_accents(s: str) -> str:
    nkfd = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in nkfd if not unicodedata.combining(c))

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", remove_accents(s).lower().strip())

def find_header_row(df: pd.DataFrame) -> int | None:
    keys = ["hoạt chất","tên thành phần","số lượng","nồng độ","hàm lượng","nhóm thuốc"]
    for i in range(min(20, len(df))):
        found = set()
        for cell in df.iloc[i].fillna(""):
            txt = str(cell).lower()
            for k in keys:
                if k in txt:
                    found.add(k)
        if len(found) >= 2:
            return i
    return None

@st.cache_data
def load_defaults():
    urls = {
        "file2":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        "file3":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        "file4":"https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    }
    f2 = pd.read_excel(BytesIO(requests.get(urls["file2"]).content))
    f3 = pd.read_excel(BytesIO(requests.get(urls["file3"]).content))
    f4 = pd.read_excel(BytesIO(requests.get(urls["file4"]).content))
    return f2, f3, f4

# ——— Load ———
file2, file3, file4 = load_defaults()

# filter file3
file3["Địa bàn"] = file3["Địa bàn"].fillna("")
file3 = file3[~file3["Địa bàn"].str.contains("tạm ngưng triển khai|ko có địa bàn", case=False)]
st.session_state["file3"] = file3

# sidebar
st.sidebar.title("Chức năng")
option = st.sidebar.radio("", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai",
])

# precalc normalized keywords for conc detection
conc_keys = [norm("nồng độ"), norm("hàm lượng")]

# ——— 1. Lọc Danh Mục Thầu ———
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")

    # chọn miền/vùng/tỉnh/bv
    df3 = st.session_state["file3"].copy()
    R = st.selectbox("Miền", ["(Tất cả)"] + sorted(df3["Miền"].dropna().unique()))
    if R!="(Tất cả)": df3 = df3[df3["Miền"]==R]
    A = st.selectbox("Vùng", ["(Tất cả)"] + sorted(df3["Vùng"].dropna().unique()))
    if A!="(Tất cả)"]: df3 = df3[df3["Vùng"]==A]
    P = st.selectbox("Tỉnh", ["(Tất cả)"] + sorted(df3["Tỉnh"].dropna().unique()))
    if P!="(Tất cả)"]: df3 = df3[df3["Tỉnh"]==P]
    H = st.selectbox("BV/SYT", sorted(df3["Bệnh viện/SYT"].dropna().unique()))

    up = st.file_uploader("File Mời Thầu (.xlsx)", type="xlsx")
    if not up:
        st.info("Tải lên file mời thầu")
        st.stop()

    xls = pd.ExcelFile(up)
    sheet = max(xls.sheet_names, key=lambda n: xls.parse(n,nrows=1,header=None).shape[1])
    raw = pd.read_excel(up, sheet_name=sheet, header=None)

    hdr = find_header_row(raw)
    if hdr is None:
        st.error("Không tìm thấy header trong 20 dòng đầu.")
        st.stop()

    cols = raw.iloc[hdr].tolist()
    df = raw.iloc[hdr+1:].reset_index(drop=True)
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)

    # --- TỰ ĐỘNG TÌM CỘT HÀM LƯỢNG ---
    conc_col = None
    for c in df.columns:
        cc = norm(c)
        if any(k in cc for k in conc_keys):
            conc_col = c
            break
    if conc_col is None:
        st.error("Không tìm thấy cột nồng độ/hàm lượng.")
        st.stop()

    # chuẩn hóa keys
    df["_act"]  = df["Tên hoạt chất"].apply(norm)
    df["_conc"] = df[conc_col].apply(norm)
    df["_grp"]  = df["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))

    # company list
    cmp = file2.copy()
    cmp["_act"]  = cmp["Tên hoạt chất"].apply(norm)
    cmp["_conc"] = cmp["Nồng độ/Hàm lượng"].apply(norm)
    cmp["_grp"]  = cmp["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))

    merged = pd.merge(df, cmp, on=["_act","_conc","_grp"], how="left", suffixes=("","_cmp"))

    info3 = df3[df3["Bệnh viện/SYT"]==H][
        ["Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai"]
    ].drop_duplicates()
    merged = pd.merge(merged, info3, on="Tên sản phẩm", how="left")

    # tỷ trọng nhóm N1–N5
    merged["Số lượng"] = pd.to_numeric(merged.get("Số lượng",0),errors="coerce").fillna(0)
    grp_ok = merged["_grp"].isin([str(i) for i in range(1,6)])
    tot = merged[grp_ok].groupby("Nhóm thuốc")["Số lượng"].transform("sum")
    merged["Tỷ trọng nhóm thầu"] = 0
    merged.loc[grp_ok,"Tỷ trọng nhóm thầu"] = (
        merged.loc[grp_ok,"Số lượng"] / tot
    ).fillna(0).apply(lambda x: f"{x:.2%}")

    st.session_state["merged"] = merged
    st.session_state["df_body"] = df

    st.success(f"✅ Đã lọc xong {len(merged)} dòng.")
    disp = merged.drop_duplicates(subset=["_act","_conc","_grp"])
    disp = disp[disp["Tên sản phẩm"].notna()]
    st.dataframe(
        disp[[
            "Tên hoạt chất", conc_col, "Nhóm thuốc",
            "Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai",
            "Tỷ trọng nhóm thầu"
        ]],
        height=500
    )

    buf = BytesIO()
    merged.to_excel(buf,index=False,sheet_name="KQ")
    st.download_button("⬇️ Tải full", buf.getvalue(), "KQ.xlsx")

# ——— 2. Phân Tích Danh Mục Thầu ———
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if "df_body" not in st.session_state:
        st.info("Chạy Lọc trước.")
        st.stop()

    dfA = st.session_state["df_body"].copy()
    dfA["Số lượng"] = pd.to_numeric(dfA["Số lượng"],errors="coerce").fillna(0)
    dfA["Giá kế hoạch"] = pd.to_numeric(dfA.get("Giá kế hoạch",0),errors="coerce").fillna(0)
    dfA["Trị giá"] = dfA["Số lượng"]*dfA["Giá kế hoạch"]

    term = st.text_input("🔍 Tra cứu hoạt chất")
    if term:
        dfA = dfA[dfA["Tên hoạt chất"].str.contains(term, case=False, na=False)]

    pd.options.display.float_format = "{:,.0f}".format
    def plot_bar(df,x,y,t):
        fig=px.bar(df,x=x,y=y,title=t); fig.update_traces(texttemplate="%{y:,.0f}",textposition="outside")
        st.plotly_chart(fig,use_container_width=True)

    g1 = dfA.groupby("Nhóm thuốc")["Trị giá"].sum().reset_index().sort_values("Trị giá",False)
    plot_bar(g1,"Nhóm thuốc","Trị giá","Trị giá theo Nhóm thuốc")
    dfA["Đường"] = dfA["Đường dùng"].apply(lambda s: "Tiêm" if "tiêm" in str(s).lower() else ("Uống" if "uống" in str(s).lower() else "Khác"))
    g2 = dfA.groupby("Đường")["Trị giá"].sum().reset_index()
    plot_bar(g2,"Đường","Trị giá","Trị giá theo Đường dùng")
    top_sl = dfA.groupby("Tên hoạt chất")["Số lượng"].sum().reset_index().sort_values("Số lượng",False).head(10)
    top_v  = dfA.groupby("Tên hoạt chất")["Trị giá"].sum().reset_index().sort_values("Trị giá",False).head(10)
    plot_bar(top_sl,"Tên hoạt chất","Số lượng","Top10 HC SL")
    plot_bar(top_v,"Tên hoạt chất","Trị giá","Top10 HC Trị giá")

    tm = {norm(a):g for a,g in zip(file4["Hoạt chất"],file4["Nhóm điều trị"])}
    dfA["Nhóm điều trị"] = dfA["Tên hoạt chất"].apply(lambda x: tm.get(norm(x),"Khác"))
    t2 = dfA.groupby("Nhóm điều trị")[["Số lượng","Trị giá"]].sum().reset_index()
    plot_bar(t2.sort_values("Số lượng",False),"Nhóm điều trị","Số lượng","SL theo Điều trị")
    sel = st.selectbox("Chọn nhóm điều trị xem Top10 HC (Trị giá)",t2["Nhóm điều trị"])
    if sel:
        t3 = dfA[dfA["Nhóm điều trị"]==sel].groupby("Tên hoạt chất")["Trị giá"].sum().reset_index().sort_values("Trị giá",False).head(10)
        plot_bar(t3,"Tên hoạt chất","Trị giá",f"Top10 HC Trị giá - {sel}")

# ——— 3. Phân Tích Danh Mục Trúng Thầu ———
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win = st.file_uploader("File Trúng Thầu (.xlsx)",type="xlsx")
    if not win:
        st.info("Tải file trúng thầu")
        st.stop()

    xlsw=pd.ExcelFile(win)
    sw=max(xlsw.sheet_names,key=lambda n:xlsw.parse(n,nrows=1,header=None).shape[1])
    raww=pd.read_excel(win,sw,header=None)
    hw=find_header_row(raww)
    if hw is None:
        st.error("Không tìm header trúng thầu")
        st.stop()

    hdrw=raww.iloc[hw].tolist()
    dfw=raww.iloc[hw+1:].reset_index(drop=True); dfw.columns=hdrw; dfw=dfw.dropna(how="all").reset_index(drop=True)
    dfw["Số lượng"]=pd.to_numeric(dfw.get("Số lượng",0),errors="coerce").fillna(0)
    pcol=next((c for c in dfw.columns if "giá trúng" in norm(c)),"Giá kế hoạch")
    dfw[pcol]=pd.to_numeric(dfw.get(pcol,0),errors="coerce").fillna(0)
    dfw["Trị giá"]=dfw["Số lượng"]*dfw[pcol]
    wv=dfw.groupby("Nhà thầu trúng")["Trị giá"].sum().reset_index().sort_values("Trị giá",False).head(20)
    f1=px.bar(wv,x="Trị giá",y="Nhà thầu trúng",orientation="h",title="Top20 Nhà thầu")
    f1.update_traces(texttemplate="%{x:,.0f}",textposition="outside"); st.plotly_chart(f1,use_container_width=True)
    dfw["Nhóm điều trị"]=dfw["Tên hoạt chất"].apply(lambda x: tm.get(norm(x),"Khác"))
    tw=dfw.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index()
    f2=px.pie(tw,names="Nhóm điều trị",values="Trị giá",title="Cơ cấu trúng thầu"); st.plotly_chart(f2,use_container_width=True)

# ——— 4. Đề Xuất Hướng Triển Khai ———
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if "merged" not in st.session_state:
        st.info("Chạy Lọc trước.")
        st.stop()

    mdf = st.session_state["merged"].copy()
    done = mdf.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Đã làm"})
    # prepare req
    req = file3.copy()
    req["_act"]=req["Tên hoạt chất"].apply(norm)
    conc3 = next((c for c in file3.columns if any(k in norm(c) for k in conc_keys)), None)
    if conc3 is None:
        st.error("Không tìm cột hàm lượng file3")
        st.stop()
    req["_conc"]=req[conc3].apply(norm)
    req["_grp"]=req["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))
    req=req.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Yêu cầu"})
    sug=pd.merge(req,done,on=["_act","_conc","_grp"],how="left").fillna(0)
    sug["Đề xuất"]=(sug["SL_Yêu cầu"]-sug["SL_Đã làm"]).clip(lower=0).astype(int)
    kh=file3.copy()
    kh["_act"]=kh["Tên hoạt chất"].apply(norm)
    kh["_conc"]=kh[conc3].apply(norm)
    kh["_grp"]=kh["Nhóm thuốc"].astype(str).apply(lambda x:re.sub(r"\D","",x))
    kh=kh.groupby(["_act","_conc","_grp"])["Tên Khách hàng phụ trách triển khai"].first().reset_index()
    sug=pd.merge(sug,kh,on=["_act","_conc","_grp"],how="left")
    st.dataframe(sug,500)
