import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
import requests
from io import BytesIO
import plotly.express as px

# --- Helper functions ---
def remove_accents(input_str):
    nfkd = unicodedata.normalize('NFKD', input_str)
    return "".join(c for c in nfkd if not unicodedata.combining(c))

def norm(s: str) -> str:
    s = remove_accents(str(s)).lower().strip()
    return re.sub(r"\s+", " ", s)

def find_header_row(df: pd.DataFrame, keywords: list):
    for i in range(10):
        row = " ".join(df.iloc[i].fillna("").astype(str)).lower()
        if all(k in row for k in keywords):
            return i
    return None

@st.cache_data
def load_default_data():
    url2 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx"
    url3 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx"
    url4 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    f2 = pd.read_excel(BytesIO(requests.get(url2).content))
    f3 = pd.read_excel(BytesIO(requests.get(url3).content))
    f4 = pd.read_excel(BytesIO(requests.get(url4).content))
    return f2, f3, f4

# Load data
file2, file3, file4 = load_default_data()

# Filter file3: loại bỏ "tạm ngưng triển khai" & "ko có địa bàn"
file3["Địa bàn"] = file3["Địa bàn"].fillna("")
file3_filtered = file3[~file3["Địa bàn"].str.contains("tạm ngưng triển khai|ko có địa bàn", case=False)]
st.session_state["file3_filtered"] = file3_filtered

# Sidebar
st.sidebar.title("Chức năng")
option = st.sidebar.radio("Chọn chức năng", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai",
    "Tra Cuu Hoat Chat"
])

# 1. Lọc Danh Mục Thầu
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    # Select region/area/province/hospital
    df3 = file3_filtered.copy()
    regions = ["(Tất cả)"] + sorted(df3["Miền"].dropna().unique().tolist())
    r = st.selectbox("Chọn Miền", regions)
    if r != "(Tất cả)": df3 = df3[df3["Miền"]==r]
    areas = ["(Tất cả)"] + sorted(df3["Vùng"].dropna().unique().tolist())
    a = st.selectbox("Chọn Vùng", areas)
    if a != "(Tất cả)": df3 = df3[df3["Vùng"]==a]
    provs = ["(Tất cả)"] + sorted(df3["Tỉnh"].dropna().unique().tolist())
    p = st.selectbox("Chọn Tỉnh", provs)
    if p != "(Tất cả)": df3 = df3[df3["Tỉnh"]==p]
    h = st.selectbox("Chọn Bệnh viện/SYT", sorted(df3["Bệnh viện/SYT"].dropna().unique().tolist()))

    uploaded = st.file_uploader("File Danh Mục Mời Thầu (.xlsx)", type="xlsx")
    if uploaded is None:
        st.info("Vui lòng tải lên file Mời Thầu.")
        st.stop()

    # Read sheet with most columns
    xls = pd.ExcelFile(uploaded)
    sheet = max(xls.sheet_names, key=lambda n: xls.parse(n, nrows=1, header=None).shape[1])
    raw = pd.read_excel(uploaded, sheet_name=sheet, header=None)
    hdr = find_header_row(raw, ["tên hoạt chất","số lượng"])
    if hdr is None:
        st.error("Không tìm thấy header trong 10 dòng đầu.")
        st.stop()

    header = raw.iloc[hdr].tolist()
    df_all = raw.iloc[hdr+1:].copy().reset_index(drop=True)
    df_all.columns = header
    df_all = df_all.dropna(how="all").reset_index(drop=True)

    # Normalize keys
    df_all["_act"] = df_all["Tên hoạt chất"].apply(norm)
    col_conc = "Nồng độ/hàm lượng" if "Nồng độ/hàm lượng" in df_all.columns else "Nồng độ/Hàm lượng"
    df_all["_conc"] = df_all[col_conc].apply(norm)
    df_all["_grp"]  = df_all["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))

    df_cmp = file2.copy()
    df_cmp["_act"]  = df_cmp["Tên hoạt chất"].apply(norm)
    df_cmp["_conc"] = df_cmp["Nồng độ/Hàm lượng"].apply(norm)
    df_cmp["_grp"]  = df_cmp["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))

    # Left merge to keep all df_all rows
    merged = pd.merge(df_all, df_cmp,
                      on=["_act","_conc","_grp"],
                      how="left", suffixes=("","_cmp"))

    # Attach branding & hospital info
    info3 = file3_filtered[file3_filtered["Bệnh viện/SYT"]==h][
        ["Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai"]].drop_duplicates()
    merged = pd.merge(merged, info3, on="Tên sản phẩm", how="left")

    # Compute Tỷ trọng nhóm thầu
    merged["Số lượng"] = pd.to_numeric(merged.get("Số lượng",0),errors="coerce").fillna(0)
    grp_sum = merged.groupby("Nhóm thuốc")["Số lượng"].transform("sum")
    merged["Tỷ trọng nhóm thầu"] = (merged["Số lượng"]/grp_sum).fillna(0).apply(lambda x:f"{x:.2%}")

    # Save for analysis
    st.session_state["filtered_df"] = merged
    st.session_state["df_all_session"] = df_all  # toàn bộ file1 data

    st.success(f"✅ Đã lọc xong {len(merged)} dòng.")
    st.dataframe(merged, height=600)

    buf = BytesIO()
    merged.to_excel(buf, index=False, sheet_name="KetQuaLoc")
    st.download_button("⬇️ Tải về kết quả", data=buf.getvalue(),file_name="KetQuaLoc.xlsx")

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if "df_all_session" not in st.session_state:
        st.info("Chạy Lọc Danh Mục Thầu trước.")
        st.stop()
    df_all = st.session_state["df_all_session"].copy()
    df_all["Số lượng"] = pd.to_numeric(df_all["Số lượng"],errors="coerce").fillna(0)
    df_all["Giá kế hoạch"] = pd.to_numeric(df_all.get("Giá kế hoạch",0),errors="coerce").fillna(0)
    df_all["Trị giá"] = df_all["Số lượng"] * df_all["Giá kế hoạch"]

    # Nhóm điều trị theo Đường dùng
    def route(x):
        s=str(x).lower()
        if "tiêm" in s: return "Tiêm"
        if "uống" in s: return "Uống"
        return "Khác"
    df_all["Đường"] = df_all.get("Đường dùng",df_all.get("Loại đường dùng","")).apply(route)

    # Chart: Trị giá theo Nhóm thuốc
    gv = df_all.groupby("Nhóm thuốc")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False)
    fig = px.bar(gv, x="Nhóm thuốc", y="Trị giá", title="Trị giá theo Nhóm thuốc")
    fig.update_traces(texttemplate="%{y:.2s}", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    # Chart: Trị giá theo Đường dùng
    gv2= df_all.groupby("Đường")["Trị giá"].sum().reset_index()
    fig2= px.bar(gv2, x="Đường", y="Trị giá", title="Trị giá theo Đường dùng")
    fig2.update_traces(texttemplate="%{y:.2s}", textposition="outside")
    st.plotly_chart(fig2, use_container_width=True)

    # Top Hoạt chất
    topA= df_all.groupby("Tên hoạt chất")["Số lượng"].sum().reset_index().sort_values("Số lượng",ascending=False).head(10)
    fig3= px.bar(topA, x="Tên hoạt chất",y="Số lượng",title="Top 10 hoạt chất (SL)")
    fig3.update_traces(texttemplate="%{y:.0f}", textposition="outside")
    st.plotly_chart(fig3,use_container_width=True)

    # Nhóm điều trị
    treat_map = {norm(a):g for a,g in zip(file4["Hoạt chất"],file4["Nhóm điều trị"])}
    df_all["Nhóm điều trị"] = df_all["Tên hoạt chất"].apply(lambda x: treat_map.get(norm(x),"Khác"))
    tv = df_all.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False)
    fig4= px.bar(tv, x="Nhóm điều trị", y="Trị giá", orientation='h',title="Trị giá theo Nhóm điều trị")
    fig4.update_traces(texttemplate="%{x:.2s}", textposition="outside")
    st.plotly_chart(fig4,use_container_width=True)

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win = st.file_uploader("File Kết quả Trúng Thầu", type="xlsx")
    if win is None:
        st.info("Hãy tải file Trúng Thầu lên.")
        st.stop()
    xlsw = pd.ExcelFile(win)
    sheetw = max(xlsw.sheet_names, key=lambda n: xlsw.parse(n,nrows=1,header=None).shape[1])
    raww = pd.read_excel(win, sheet_name=sheetw, header=None)
    hi = find_header_row(raww, ["tên hoạt chất","nhà thầu trúng"])
    if hi is None:
        st.error("Không tìm header trúng thầu.")
        st.stop()
    hdrw = raww.iloc[hi].tolist()
    dfw = raww.iloc[hi+1:].copy().reset_index(drop=True)
    dfw.columns=hdrw
    dfw=dfw.dropna(how="all").reset_index(drop=True)
    dfw["Số lượng"] = pd.to_numeric(dfw.get("Số lượng",0),errors="coerce").fillna(0)
    price_col = next((c for c in dfw.columns if "Giá trúng" in c), "Giá kế hoạch")
    dfw[price_col]=pd.to_numeric(dfw.get(price_col,0),errors="coerce").fillna(0)
    dfw["Trị giá"]=dfw["Số lượng"]*dfw[price_col]

    wv = dfw.groupby("Nhà thầu trúng")["Trị giá"].sum().reset_index().sort_values("Trị giá",ascending=False).head(20)
    f1=px.bar(wv, x="Trị giá",y="Nhà thầu trúng",orientation='h',title="Top 20 nhà thầu trúng")
    f1.update_traces(texttemplate="%{x:.2s}",textposition="outside")
    st.plotly_chart(f1,use_container_width=True)

    dfw["Nhóm điều trị"] = dfw["Tên hoạt chất"].apply(lambda x: treat_map.get(norm(x),"Khác"))
    tw = dfw.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index()
    f2=px.pie(tw,names="Nhóm điều trị",values="Trị giá",title="Cơ cấu trị giá trúng thầu")
    st.plotly_chart(f2,use_container_width=True)

# 4. Đề Xuất Hướng Triển Khai
elif option == "Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if "filtered_df" not in st.session_state:
        st.info("Chạy Lọc Danh Mục Thầu trước.")
        st.stop()
    dfm = st.session_state["filtered_df"]
    # Tổng SL đã làm theo HC/HML/NT
    sl_done = dfm.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_Đã làm"})
    # Tổng SL BV cần (từ file3_filtered)
    sl_req = file3_filtered.copy()
    sl_req["_act"]=sl_req["Tên hoạt chất"].apply(norm)
    sl_req["_conc"]=sl_req[col_conc].apply(norm)
    sl_req["_grp"]=sl_req["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))
    sl_req = sl_req.groupby(["_act","_conc","_grp"])["Số lượng"].sum().reset_index().rename(columns={"Số lượng":"SL_BV"})
    # Ghép
    sug = pd.merge(sl_req, sl_done, on=["_act","_conc","_grp"], how="left")
    sug["SL_Đã làm"] = sug["SL_Đã làm"].fillna(0).astype(int)
    sug["Đề xuất"]= (sug["SL_BV"] - sug["SL_Đã làm"]).clip(lower=0).astype(int)
    # Thêm Khách hàng
    kh = file3_filtered[["_act","_conc","_grp","Tên Khách hàng phụ trách triển khai"]]
    kh["_act"]=kh["Tên hoạt chất"].apply(norm)
    kh[col_conc]=kh[col_conc].apply(norm)
    kh["_grp"]=kh["Nhóm thuốc"].astype(str).apply(lambda x: re.sub(r"\D","",x))
    kh = kh.groupby(["_act","_conc","_grp"])["Tên Khách hàng phụ trách triển khai"].first().reset_index()
    sug = sug.merge(kh, on=["_act","_conc","_grp"], how="left")

    st.subheader("📦 Bảng đề xuất cơ số thầu tới")
    st.dataframe(sug, height=500)

# 5. Tra cứu hoạt chất
elif option == "Tra Cuu Hoat Chat":
    st.header("🔍 Tra cứu hoạt chất")
    term = st.text_input("Nhập hoạt chất")
    if term:
        out = file4[file4["Hoạt chất"].str.contains(term, case=False, na=False)]
        if out.empty:
            st.warning("Không tìm thấy.")
        else:
            st.dataframe(out)
