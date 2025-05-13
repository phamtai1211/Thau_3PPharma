import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import requests
from io import BytesIO
import plotly.express as px

# —— Helper Functions ——
def remove_accents(s: str) -> str:
    nkfd = unicodedata.normalize("NFKD", str(s))
    return "".join(c for c in nkfd if not unicodedata.combining(c))

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", remove_accents(s).lower().strip())

def find_header_row(df: pd.DataFrame) -> int | None:
    keys = ["hoạt chất","tên thành phần","số lượng","nồng độ","hàm lượng","nhóm thuốc"]
    for i in range(min(20, len(df))):
        text = " ".join(df.iloc[i].fillna("").astype(str).tolist()).lower()
        if any(k in text for k in ["hoạt chất"]) and any(k in text for k in ["số lượng","nồng độ"]):
            return i
    return None

@st.cache_data
def load_defaults():
    urls = {
        "file2": "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx",
        "file3": "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx",
        "file4": "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    }
    f2 = pd.read_excel(BytesIO(requests.get(urls["file2"]).content))
    f3 = pd.read_excel(BytesIO(requests.get(urls["file3"]).content))
    f4 = pd.read_excel(BytesIO(requests.get(urls["file4"]).content))
    return f2, f3, f4

# —— Load Data ——
file2, file3, file4 = load_defaults()

# Filter inactive areas
file3["Địa bàn"] = file3["Địa bàn"].fillna("")
file3 = file3[~file3["Địa bàn"].str.contains("tạm ngưng triển khai|ko có địa bàn", case=False)]
st.session_state["file3"] = file3

# Sidebar
st.sidebar.title("Chức năng")
option = st.sidebar.radio("", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai",
])
conc_keys = ["nồng độ","hàm lượng"]

# —— 1. Lọc Danh Mục Thầu ——
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3 = st.session_state["file3"].copy()
    R = st.selectbox("Miền", ["(Tất cả)"] + sorted(df3["Miền"].dropna().unique()))
    if R != "(Tất cả)": df3 = df3[df3["Miền"] == R]
    A = st.selectbox("Vùng", ["(Tất cả)"] + sorted(df3["Vùng"].dropna().unique()))
    if A != "(Tất cả)": df3 = df3[df3["Vùng"] == A]
    P = st.selectbox("Tỉnh", ["(Tất cả)"] + sorted(df3["Tỉnh"].dropna().unique()))
    if P != "(Tất cả)": df3 = df3[df3["Tỉnh"] == P]
    H = st.selectbox("BV/SYT", sorted(df3["Bệnh viện/SYT"].dropna().unique()))

    up = st.file_uploader("File Mời Thầu (.xlsx)", type="xlsx")
    if not up:
        st.info("Tải lên file mời thầu")
        st.stop()

    xls = pd.ExcelFile(up)
    sheet = max(xls.sheet_names, key=lambda n: xls.parse(n, nrows=1, header=None).shape[1])
    raw = pd.read_excel(up, sheet, header=None)
    hdr = find_header_row(raw)
    if hdr is None:
        st.error("Không tìm thấy header trong 20 dòng đầu.")
        st.stop()

    cols = raw.iloc[hdr].tolist()
    df = raw.iloc[hdr+1:].reset_index(drop=True)
    df.columns = cols
    df = df.dropna(how="all").reset_index(drop=True)

    # Dynamic column detection
    act_col = next((c for c in df.columns if "hoạt chất" in norm(c) or "tên thành phần" in norm(c)), None)
    conc_col = next((c for c in df.columns if any(k in norm(c) for k in conc_keys)), None)
    grp_col = next((c for c in df.columns if "nhóm" in norm(c)), "Nhóm thuốc")
    if not act_col or not conc_col:
        st.error("Không tìm thấy cột hoạt chất hoặc hàm lượng.")
        st.stop()

    df["_act"] = df[act_col].apply(norm)
    df["_conc"] = df[conc_col].apply(norm)
    df["_grp"] = df[grp_col].astype(str).apply(lambda x: re.sub(r"\D","",x))

    cmp = file2.copy()
    cmp_act = next((c for c in cmp.columns if "hoạt chất" in norm(c)), "Tên hoạt chất")
    cmp_conc = next((c for c in cmp.columns if any(k in norm(c) for k in conc_keys)), "Nồng độ/Hàm lượng")
    cmp_grp = next((c for c in cmp.columns if "nhóm" in norm(c)), "Nhóm thuốc")
    cmp["_act"] = cmp[cmp_act].apply(norm)
    cmp["_conc"] = cmp[cmp_conc].apply(norm)
    cmp["_grp"] = cmp[cmp_grp].astype(str).apply(lambda x: re.sub(r"\D","",x))

    merged = pd.merge(df, cmp, on=["_act","_conc","_grp"], how="left", suffixes=("","_cmp"))
    info3 = df3[df3["Bệnh viện/SYT"]==H][["Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai"]].drop_duplicates()
    merged = pd.merge(merged, info3, on="Tên sản phẩm", how="left")

    merged["Số lượng"] = pd.to_numeric(merged.get("Số lượng",0), errors="coerce").fillna(0)
    valid = merged["_grp"].isin([str(i) for i in range(1,6)])
    grp_sum = merged[valid].groupby("Nhóm thuốc")["Số lượng"].transform("sum")
    merged["Tỷ trọng nhóm thầu"] = 0
    merged.loc[valid,"Tỷ trọng nhóm thầu"] = (merged.loc[valid,"Số lượng"]/grp_sum).apply(lambda x: f"{x:.2%}")

    st.success(f"✅ Đã lọc xong {len(merged)} dòng.")
    disp = merged.drop_duplicates(subset=["_act","_conc","_grp"])
    st.dataframe(disp[[act_col,conc_col,grp_col,"Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai","Tỷ trọng nhóm thầu"]], height=500)

    buf = BytesIO()
    merged.to_excel(buf, index=False, sheet_name="KếtQuả")
    st.download_button("⬇️ Tải full", buf.getvalue(), "KetQuaLoc.xlsx")
    st.session_state.update({"merged":merged,"df_body":df})

# —— 2. Phân Tích Danh Mục Thầu ——
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if "df_body" not in st.session_state:
        st.info("Chạy Lọc Danh Mục Thầu trước.")
        st.stop()

    dfA = st.session_state["df_body"].copy()
    dfA["Số lượng"] = pd.to_numeric(dfA.get("Số lượng",0), errors="coerce").fillna(0)
    dfA["Giá kế hoạch"] = pd.to_numeric(dfA.get("Giá kế hoạch",0), errors="coerce").fillna(0)
    dfA["Trị giá"] = dfA["Số lượng"]*dfA["Giá kế hoạch"]

    term = st.text_input("🔍 Tra cứu hoạt chất (nhập một phần)")
    if term:
        dfA = dfA[dfA[act_col].str.contains(term, case=False, na=False)]

    def plot(df, x, y, title):
        fig = px.bar(df, x=x, y=y, title=title)
        fig.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    # Nhóm thuốc tổng
    plot(dfA.groupby(grp_col)["Trị giá"].sum().reset_index(), grp_col, "Trị giá", "Trị giá theo Nhóm thuốc")
    # Đường dùng
    dfA["Đường"] = dfA["Đường dùng"].apply(lambda s: "Tiêm" if "tiêm" in str(s).lower() else ("Uống" if "uống" in str(s).lower() else "Khác"))
    plot(dfA.groupby("Đường")["Trị giá"].sum().reset_index(), "Đường","Trị giá","Trị giá theo Đường dùng")

    # Top10 HC phân theo Tiêm/Uống và SL/Trị giá
    for route in ["Tiêm","Uống"]:
        sub = dfA[dfA["Đường"]==route]
        plot(sub.groupby(act_col)["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False).head(10), act_col, "Trị giá", f"Top10 {route} theo Trị giá")
        plot(sub.groupby(act_col)["Số lượng"].sum().reset_index().sort_values("Số lượng",ascending=False).head(10), act_col, "Số lượng", f"Top10 {route} theo Số lượng")

    # Nhóm điều trị
    tm = {norm(a):g for a,g in zip(file4["Hoạt chất"],file4["Nhóm điều trị"])}
    dfA["Nhóm điều trị"] = dfA[act_col].apply(lambda x: tm.get(norm(x),"Khác"))
    t2 = dfA.groupby("Nhóm điều trị")[ ["Số lượng","Trị giá"] ].sum().reset_index()
    plot(t2.sort_values("Số lượng",False),"Nhóm điều trị","Số lượng","SL theo Nhóm điều trị")
    sel = st.selectbox("Chọn Nhóm điều trị xem Top10 HC (Trị giá)", t2["Nhóm điều trị"])
    if sel:
        plot(dfA[dfA["Nhóm điều trị"]==sel].groupby(act_col)["Trị giá"].sum().reset_index().sort_values("Trị giá",False).head(10),act_col,"Trị giá",f"Top10 HC theo Trị giá - {sel}")

# —— 3. Phân Tích Danh Mục Trúng Thầu ——
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win = st.file_uploader("File Trúng Thầu (.xlsx)", type="xlsx")
    if not win:
        st.info("Tải lên file trúng thầu")
        st.stop()
    xlsw = pd.ExcelFile(win)
    sw = max(xlsw.sheet_names, key=lambda n: xlsw.parse(n,nrows=1,header=None).shape[1])
    raww = pd.read_excel(win,sw,header=None)
    hw = find_header_row(raww)
    if hw is None:
        st.error("Không tìm header trúng thầu.")
        st.stop()
    hdrw = raww.iloc[hw].tolist()
    dfw = raww.iloc[hw+1:].reset_index(drop=True)
    dfw.columns = hdrw
    dfw = dfw.dropna(how="all").reset_index(drop=True)
    dfw["Số lượng"] = pd.to_numeric(dfw.get("Số lượng",0),errors="coerce").fillna(0)
    pcol = next((c for c in dfw.columns if "giá trúng" in norm(c)), "Giá kế hoạch")
    dfw[pcol] = pd.to_numeric(dfw.get(pcol,0),errors="coerce").fillna(0)
    dfw["Trị giá"] = dfw["Số lượng"]*dfw[pcol]
    wv = dfw.groupby("Nhà thầu trúng")["Trị giá"].sum().reset_index().sort_values("Trị giá",False).head(20)
    f1=px.bar(wv,x="Trị giá",y="Nhà thầu trúng",orientation="h",title="Top20 Nhà thầu trúng")
    f1.update_traces(texttemplate="%{x:,.0f}",textposition="outside")
    st.plotly_chart(f1,use_container_width=True)
    dfw["Nhóm điều trị"] = dfw[act_col].apply(lambda x: tm.get(norm(x),"Khác"))
    tw = dfw.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index()
    f2=px.pie(tw,names="Nhóm điều trị",values="Trị giá",title="Cơ cấu trúng thầu")
    st.plotly_chart(f2,use_container_width=True)

# —— 4. Đề Xuất Hướng Triển Khai ——
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if "merged" not in st.session_state:
        st.info("Chạy Lọc trước.")
        st.stop()
    mdf = st.session_state["merged"].copy()
    done = mdf.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Đã làm"})
    req = file3.copy()
    act3 = next((c for c in req.columns if "hoạt chất" in norm(c)), None)
    conc3 = next((c for c in req.columns if any(k in norm(c) for k in conc_keys)), None)
    grp3  = next((c for c in req.columns if "nhóm" in norm(c)), None)
    if not act3 or not conc3 or not grp3:
        st.error("Không tìm đủ cột trong file3 để đề xuất.")
        st.stop()
    req["_act"] = req[act3].apply(norm)
    req["_conc"] = req[conc3].apply(norm)
    req["_grp"] = req[grp3].astype(str).apply(lambda x: re.sub(r"\D","",x))
    req = req.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Yêu cầu"})
    sug = pd.merge(req, done, on=["_act","_conc","_grp"], how="left").fillna(0)
    sug["Đề xuất"] = (sug["SL_Yêu cầu"]-sug["SL_Đã làm"]).clip(lower=0).astype(int)
    kh = file3.copy()
    kh["_act"] = kh[act3].apply(norm)
    kh["_conc"] = kh[conc3].apply(norm)
    kh["_grp"] = kh[grp3].astype(str).apply(lambda x: re.sub(r"\D","",x))
    kh = kh.groupby(["_act","_conc","_grp"])["Tên Khách hàng phụ trách triển khai"].first().reset_index()
    sug = pd.merge(sug, kh, on=["_act","_conc","_grp"], how="left")
    st.dataframe(sug.sort_values("Đề xuất",False).reset_index(drop=True))
