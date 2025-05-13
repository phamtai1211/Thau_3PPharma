import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO
import plotly.express as px

# --- Hàm dò header ---
def find_header_index(df_raw):
    for i in range(10):
        row = " ".join(df_raw.iloc[i].fillna('').astype(str).tolist()).lower()
        if "tên hoạt chất" in row and "số lượng" in row:
            return i
    return None

# --- Chuẩn hóa ---
def normalize_active(name: str) -> str:
    return re.sub(r'\s+', ' ',
                  re.sub(r'\(.*?\)', '',
                         str(name))).strip().lower()

def normalize_concentration(conc: str) -> str:
    s = str(conc).lower().replace(',', '.').replace('dung tích', '')
    parts = [p.strip() for p in s.split(',') if re.search(r'\d', p)]
    if len(parts)>=2 and re.search(r'(mg|mcg|g|%)', parts[0]) and 'ml' in parts[-1] and '/' not in parts[0]:
        return parts[0].replace(' ','') + '/' + parts[-1].replace(' ','')
    return ''.join(p.replace(' ','') for p in parts)

def normalize_group(grp: str) -> str:
    return re.sub(r'\D', '', str(grp)).strip()

# --- Load dữ liệu công ty & BV ---
@st.cache_data
def load_default_data():
    url2 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file2.xlsx"
    url3 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/file3.xlsx"
    url4 = "https://raw.githubusercontent.com/phamtai1211/Thau_3PPharma/main/nhom_dieu_tri.xlsx"
    f2 = pd.read_excel(BytesIO(requests.get(url2).content))
    f3 = pd.read_excel(BytesIO(requests.get(url3).content))
    f4 = pd.read_excel(BytesIO(requests.get(url4).content))
    return f2, f3, f4

file2, file3, file4 = load_default_data()

# Lọc file3 (loại bỏ tạm ngưng/ko có địa bàn)
file3 = file3[~file3["Địa bàn"].astype(str)
              .str.contains("tạm ngưng triển khai|ko có địa bàn", case=False, na=False)]

# Sidebar
st.sidebar.title("Chức năng")
opt = st.sidebar.radio("", [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu",
    "Đề Xuất Hướng Triển Khai",
])

# 1) Lọc Danh Mục Thầu
if opt=="Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    # chọn BV
    regions = sorted(file3["Miền"].dropna().unique())
    r = st.selectbox("Miền", regions)
    sub = file3[file3["Miền"]==r]
    areas = sorted(sub["Vùng"].dropna().unique())
    a = st.selectbox("Vùng", ["(Tất cả)"]+areas)
    if a!="(Tất cả)": sub = sub[sub["Vùng"]==a]
    ps = sorted(sub["Tỉnh"].dropna().unique())
    p = st.selectbox("Tỉnh", ps)
    sub = sub[sub["Tỉnh"]==p]
    hosp = st.selectbox("BV/SYT", sorted(sub["Bệnh viện/SYT"].dropna().unique()))

    up = st.file_uploader("File Mời Thầu (.xlsx)", type="xlsx")
    if up and hosp:
        xls = pd.ExcelFile(up)
        # chọn sheet nhiều cột nhất
        best, sheet = 0, None
        for nm in xls.sheet_names:
            try:
                c = xls.parse(nm, nrows=1, header=None).shape[1]
                if c>best:
                    best, sheet = c, nm
            except: pass

        df_raw = pd.read_excel(up, sheet_name=sheet, header=None)
        hi = find_header_index(df_raw)
        if hi is None:
            st.error("❌ Không tìm thấy header trong 10 dòng đầu.")
            st.stop()

        # DataFrame gốc (sau header), **không** dropna
        df_all = df_raw.iloc[hi+1:].reset_index(drop=True)
        df_all.columns = df_raw.iloc[hi].tolist()

        # Chuẩn hóa thêm 3 cột để merge
        df_all["active_norm"] = df_all["Tên hoạt chất"].apply(normalize_active)
        df_all["conc_norm"]   = df_all["Nồng độ/hàm lượng"].apply(normalize_concentration)
        df_all["grp_norm"]    = df_all["Nhóm thuốc"].apply(normalize_group)

        # Bảng tham chiếu 1-1 từ file2
        df2 = file2.copy()
        df2["active_norm"] = df2["Tên hoạt chất"].apply(normalize_active)
        df2["conc_norm"]   = df2["Nồng độ/Hàm lượng"].apply(normalize_concentration)
        df2["grp_norm"]    = df2["Nhóm thuốc"].apply(normalize_group)
        comp_ref = df2[["active_norm","conc_norm","grp_norm","Tên sản phẩm"]].drop_duplicates(
            subset=["active_norm","conc_norm","grp_norm"]
        )

        # In số dòng gốc và merge
        n_in = df_all.shape[0]
        st.write(f"❓ Dòng sau header: **{n_in}**")
        m = pd.merge(
            df_all.reset_index(),
            comp_ref,
            on=["active_norm","conc_norm","grp_norm"],
            how="left"
        )
        res = (
            m.set_index("index")
             [ df_all.columns.tolist() + ["Tên sản phẩm"] ]
             .reset_index(drop=True)
        )
        n_out = res.shape[0]
        st.write(f"✅ Dòng sau merge: **{n_out}**")

        # Thêm Địa bàn & Khách hàng phụ trách từ file3
        hosp_df = file3[file3["Bệnh viện/SYT"]==hosp][
            ["Tên sản phẩm","Địa bàn","Tên Khách hàng phụ trách triển khai"]
        ]
        res = res.merge(hosp_df, on="Tên sản phẩm", how="left")

        # Tính tỷ trọng SL/DM Tổng theo NHÓM ĐIỀU TRỊ
        treat_map = {normalize_active(a):g for a,g in zip(
            file4["Hoạt chất"], file4["Nhóm điều trị"]
        )}
        # tổng theo nhóm trên toàn bộ mời thầu
        grp_tot = {}
        for _,r0 in df_all.iterrows():
            a0 = normalize_active(r0["Tên hoạt chất"])
            g0 = treat_map.get(a0)
            q0 = pd.to_numeric(r0.get("Số lượng",0), errors="coerce") or 0
            if g0: grp_tot[g0] = grp_tot.get(g0,0)+q0

        def calc_ratio(row):
            a0 = normalize_active(row["Tên hoạt chất"])
            g0 = treat_map.get(a0)
            q0 = pd.to_numeric(row.get("Số lượng",0), errors="coerce") or 0
            if not(g0 and grp_tot.get(g0)>0): return None
            return f"{q0/grp_tot[g0]:.2%}"

        res["Tỷ trọng SL/DM Tổng"] = res.apply(calc_ratio, axis=1)

        # Hiển thị và download
        st.dataframe(res, height=600)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            res.to_excel(w, index=False, sheet_name="KetQuaLoc")
        st.download_button("⬇️ Tải về Excel", data=buf.getvalue(),
                           file_name="Ketqua_loc.xlsx")

        # Lưu session
        st.session_state["filtered_df"] = res
        st.session_state["hosp"] = hosp

# 2) Phân Tích Danh Mục Thầu
elif opt=="Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if "filtered_df" not in st.session_state:
        st.info("Chạy 'Lọc Danh Mục Thầu' trước đã.")
        st.stop()
    df = st.session_state["filtered_df"].copy()
    df["Số lượng"] = pd.to_numeric(df["Số lượng"], errors="coerce").fillna(0)
    df["Giá kế hoạch"] = pd.to_numeric(df["Giá kế hoạch"], errors="coerce").fillna(0)
    df["Trị giá"] = df["Số lượng"] * df["Giá kế hoạch"]

    # Nhóm thầu theo trị giá
    gv = df.groupby("Nhóm thuốc")["Trị giá"].sum().reset_index().sort_values("Trị giá",False)
    fig = px.bar(gv, x="Nhóm thuốc", y="Trị giá", text="Trị giá",
                 title="Trị giá theo Nhóm thầu")
    fig.update_traces(texttemplate="%{text:.2s}", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    # Đường dùng
    def cls(rt):
        r = str(rt).lower()
        if "tiêm" in r: return "Tiêm"
        if "uống" in r: return "Uống"
        return "Khác"
    df["Đường"] = df["Đường dùng"].apply(cls)
    dv = df.groupby("Đường")["Trị giá"].sum().reset_index()
    st.plotly_chart(px.pie(dv,names="Đường",values="Trị giá",
                           title="Cơ cấu đường dùng"))

    # Top10 HC theo Trị giá/Số lượng, Tiêm/Uống
    for measure in ["Số lượng","Trị giá"]:
        for route in ["Tiêm","Uống"]:
            sub = df[df["Đường"]==route]
            top = sub.groupby("Tên hoạt chất")[measure].sum().reset_index() \
                     .sort_values(measure,False).head(10)
            st.subheader(f"Top10 HC {route} theo {measure}")
            st.plotly_chart(px.bar(top, x="Tên hoạt chất", y=measure,
                                   text=measure).update_traces(
                                   texttemplate="%{text:.2s}", textposition="outside"
            ), use_container_width=True)

    # Nhóm điều trị trị giá và SL
    tm = {normalize_active(a):g for a,g in zip(
        file4["Hoạt chất"], file4["Nhóm điều trị"]
    )}
    df["Nhóm điều trị"] = df["Tên hoạt chất"].apply(
        lambda x: tm.get(normalize_active(x),"Khác"))
    # Trị giá
    tv = df.groupby("Nhóm điều trị")["Trị giá"].sum().reset_index().sort_values("Trị giá",False)
    st.plotly_chart(px.bar(tv, x="Nhóm điều trị", y="Trị giá",
                           orientation="h", title="Trị giá theo Nhóm điều trị"),
                   use_container_width=True)
    # Số lượng
    sv = df.groupby("Nhóm điều trị")["Số lượng"].sum().reset_index().sort_values("Số lượng",False)
    st.plotly_chart(px.bar(sv, x="Nhóm điều trị", y="Số lượng",
                           orientation="h", title="SL theo Nhóm điều trị"),
                   use_container_width=True)

    # Top10 HC trong nhóm điều trị
    sel = st.selectbox("Chọn Nhóm điều trị", tv["Nhóm điều trị"])
    if sel:
        tmp = df[df["Nhóm điều trị"]==sel]
        topv = tmp.groupby("Tên hoạt chất")["Trị giá"].sum().reset_index() \
                  .sort_values("Trị giá",False).head(10)
        st.plotly_chart(px.bar(topv, x="Tên hoạt chất", y="Trị giá",
                               orientation="h",
                               title=f"Top10 HC - {sel} (Trị giá)"),
                       use_container_width=True)

# 3) Phân Tích Danh Mục Trúng Thầu
elif opt=="Phân Tích Danh Mục Trúng Thầu":
    st.header("🏆 Phân Tích Danh Mục Trúng Thầu")
    win = st.file_uploader("File Trúng Thầu", type="xlsx")
    inv = st.file_uploader("Đối chiếu Mời Thầu (tuỳ chọn)", type="xlsx")
    if not win:
        st.info("Upload file Trúng Thầu đã.")
        st.stop()
    # --- tương tự như cũ, dò header, parse, tính Trị giá, vẽ chart Nhà thầu, Nhóm điều trị ---
    # ... giữ nguyên logic ban đầu của anh ...

# 4) Đề Xuất Hướng Triển Khai
elif opt=="Đề Xuất Hướng Triển Khai":
    st.header("💡 Đề Xuất Hướng Triển Khai")
    if "filtered_df" not in st.session_state:
        st.info("Chạy phân tích trước đã.")
        st.stop()
    df = st.session_state["filtered_df"]
    hosp = st.session_state["hosp"]
    # --- logic đề xuất theo file3 và missing_items ---
    # ... giữ nguyên logic ban đầu của anh ...

