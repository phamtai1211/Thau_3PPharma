import streamlit as st
import pandas as pd
import io
import zipfile

# --- Helper functions ---
@st.cache_data
def read_excel_file(uploaded):
    """
    Đọc file Excel, tự động phát hiện dòng header nằm trong 10 dòng đầu.
    """
    df0 = pd.read_excel(uploaded, header=None)
    header_row = 0
    for i in range(min(10, len(df0))):
        row = df0.iloc[i].astype(str)
        if any("Bệnh viện" in c or "Danh Mục" in c for c in row):
            header_row = i
            break
    return pd.read_excel(uploaded, header=header_row)


def process_uploaded(uploaded, df3_temp):
    """
    Xử lý file Danh Mục Mời Thầu:
    - Đọc file
    - Lọc các dòng tồn tại trong df3_temp (dựa trên cột 'Bệnh viện/SYT')
    """
    df = read_excel_file(uploaded)
    # Giả sử file Excel có cột 'Bệnh viện/SYT'
    display_df = df[df['Bệnh viện/SYT'].isin(df3_temp['Bệnh viện/SYT'])]
    export_df = display_df.copy()
    return display_df, export_df


def to_excel_bytes(df_):
    """Chuyển DataFrame thành bytes để download Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_.to_excel(writer, index=False)
    return output.getvalue()

# --- Load reference files ---
st.sidebar.header("🔧 Tải các file tham khảo")
file3 = st.sidebar.file_uploader(
    "File 3: Danh sách triển khai (Miền, Vùng, Tỉnh, BV/SYT...)",
    type=['xlsx'], key="file3"
)
file4 = st.sidebar.file_uploader(
    "File 4: Danh sách Hoạt chất – Nhóm điều trị", type=['xlsx'], key="file4"
)

if not file3 or not file4:
    st.sidebar.warning("Vui lòng upload đủ cả File 3 và File 4 ở trên.")
    st.stop()

# Đọc file tham khảo
df3_ref = pd.read_excel(file3)
df4_ref = pd.read_excel(file4)

# --- Main UI ---
st.title("🏥 Ứng dụng Phân tích Đấu thầu Thuốc")
menu = [
    "Lọc Danh Mục Thầu",
    "Phân Tích Danh Mục Thầu",
    "Phân Tích Danh Mục Trúng Thầu"
]
option = st.sidebar.selectbox("Chọn chức năng", menu)

# 1. Lọc Danh Mục Thầu
if option == "Lọc Danh Mục Thầu":
    st.header("📂 Lọc Danh Mục Thầu")
    df3_temp = df3_ref.copy()
    for col in ['Miền','Vùng','Tỉnh','Bệnh viện/SYT']:
        opts = ['(Tất cả)'] + sorted(df3_temp[col].dropna().unique())
        sel = st.selectbox(f"Chọn {col}", opts, key=col)
        if sel != '(Tất cả)':
            df3_temp = df3_temp[df3_temp[col] == sel]

    uploaded = st.file_uploader(
        "Tải lên file Danh Mục Mời Thầu (.xlsx)", type=['xlsx']
    )
    if uploaded:
        display_df, export_df = process_uploaded(uploaded, df3_temp)
        st.success(f"✅ Tổng dòng khớp: {len(display_df)}")

        # Hiển thị bảng gốc (style giống cũ)
        display_ui = display_df.fillna('').astype(str)
        st.write(display_ui)

        # Lưu session để dùng phía sau
        st.session_state['filtered_display'] = display_df.copy()
        st.session_state['filtered_export']  = export_df.copy()
        st.session_state['file3_temp']      = df3_temp.copy()

        # --- Tính toán số liệu ---
        df_calc = display_df.copy()
        df_calc.columns = df_calc.columns.str.strip()
        df_calc['Số lượng']     = pd.to_numeric(
            df_calc.get('Số lượng', 0), errors='coerce'
        ).fillna(0)
        df_calc['Giá kế hoạch'] = pd.to_numeric(
            df_calc.get('Giá kế hoạch', 0), errors='coerce'
        ).fillna(0)
        df_calc['Trị giá']      = df_calc['Số lượng'] * df_calc['Giá kế hoạch']

        # Hàm format hiển thị
        def fmt(x):
            if x >= 1e9: return f"{x/1e9:.2f} tỷ"
            if x >= 1e6: return f"{x/1e6:.2f} triệu"
            if x >= 1e3: return f"{x/1e3:.2f} nghìn"
            return str(int(x))

        # Tổng Trị giá theo Hoạt chất
        if 'Tên hoạt chất' in df_calc.columns:
            val = (
                df_calc
                .groupby('Tên hoạt chất')['Trị giá']
                .sum()
                .reset_index()
                .sort_values('Trị giá', ascending=False)
            )
            val['Trị giá'] = val['Trị giá'].apply(fmt)
            st.subheader('Tổng Trị giá theo Hoạt chất')
            st.table(val)

            # Nút download kết quả
            excel_data = to_excel_bytes(val)
            st.download_button(
                label="📥 Tải kết quả tổng Trị giá (.xlsx)",
                data=excel_data,
                file_name="tong_tri_gia_theo_hoatchat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Không tìm thấy cột 'Tên hoạt chất'.")

# 2. Phân Tích Danh Mục Thầu
elif option == "Phân Tích Danh Mục Thầu":
    st.header("📊 Phân Tích Danh Mục Thầu")
    if 'filtered_export' in st.session_state:
        df_exp = st.session_state['filtered_export']
        file3_temp = st.session_state['file3_temp']
        summary = (
            df_exp
            .groupby(['Bệnh viện/SYT','Tên hoạt chất'])
            .agg(
                SL=('Số lượng','sum'),
                TG=('Trị giá','sum')
            )
            .reset_index()
        )
        st.subheader("Tổng SL & Trị giá theo BV/SYT – Hoạt chất")
        st.dataframe(summary)

        excel_data = to_excel_bytes(summary)
        st.download_button(
            label="📥 Tải kết quả phân tích (.xlsx)",
            data=excel_data,
            file_name="phan_tich_danh_muc_thau.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ Bạn phải chạy “Lọc Danh Mục Thầu” trước.")

# 3. Phân Tích Danh Mục Trúng Thầu
elif option == "Phân Tích Danh Mục Trúng Thầu":
    st.header("🔍 Phân Tích Danh Mục Trúng Thầu")
    st.info("Chức năng đang được xây dựng..."

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
