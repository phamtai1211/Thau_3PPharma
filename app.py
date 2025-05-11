import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
import plotly.express as px

# Cấu hình trang
st.set_page_config(page_title="Phân Tích Dữ Liệu Thầu Thuốc", layout="wide")

# Hàm tiện ích
def remove_accents(s: str) -> str:
    """Loại bỏ dấu tiếng Việt (kể cả chữ đ) khỏi chuỗi."""
    s = str(s)
    s = s.replace("đ", "d").replace("Đ", "D")
    return ''.join(ch for ch in unicodedata.normalize('NFD', s) if unicodedata.category(ch) != 'Mn')

def normalize_text(s: str) -> str:
    """Chuẩn hóa chuỗi để so sánh: hạ chữ thường, bỏ dấu, xóa khoảng trắng thừa, bỏ nội dung trong ngoặc, chuẩn hóa dấu + và /."""
    if s is None:
        return ""
    s = remove_accents(s).lower()
    s = re.sub(r'\([^)]*\)', '', s)        # bỏ nội dung trong ngoặc đơn
    s = re.sub(r'\s+', ' ', s).strip()     # gộp nhiều khoảng trắng thành một, bỏ khoảng trắng đầu/cuối
    s = re.sub(r'\s*\+\s*', '+', s)        # bỏ khoảng trắng quanh dấu +
    s = re.sub(r'\s*\/\s*', '/', s)        # bỏ khoảng trắng quanh dấu /
    s = s.replace(" ", "")                # xóa mọi khoảng trắng còn lại
    return s

def detect_header_row(file, sheet_name=0) -> int:
    """Tìm chỉ số dòng tiêu đề cột (trong 10 dòng đầu) dựa trên các từ khóa nhận dạng cột."""
    try:
        sample_df = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=10)
    except Exception:
        return None
    # Các từ khóa để nhận dạng dòng tiêu đề
    keywords = ["tên thuốc", "hoạt chất", "hàm lượng", "nồng độ", "dạng bào chế",
                "đường dùng", "số lượng", "đơn vị", "giá kế hoạch", "đơn giá trúng thầu",
                "thành tiền", "nhà thầu", "miền", "vùng", "tỉnh", "bệnh viện", "syt",
                "nhóm thuốc", "tên sản phẩm", "khách hàng"]
    keywords_norm = [remove_accents(k).lower() for k in keywords]
    header_idx = None
    for i in range(min(10, len(sample_df))):
        row_text = remove_accents(" ".join(sample_df.loc[i].fillna("").astype(str).tolist())).lower()
        count = 0
        for kw in keywords_norm:
            if kw in row_text:
                count += 1
        if count >= 3:  # dòng có từ 3 từ khóa trở lên
            header_idx = i
            break
    return header_idx

def load_excel(file):
    """Đọc file Excel vào DataFrame, tự động phát hiện dòng tiêu đề."""
    header_idx = detect_header_row(file)
    try:
        df = pd.read_excel(file, header=header_idx)
    except Exception:
        # Thử lại với engine openpyxl nếu cần
        df = pd.read_excel(file, header=header_idx, engine='openpyxl')
    return df, header_idx

def identify_columns(df: pd.DataFrame):
    """Xác định các cột quan trọng trong DataFrame dựa vào tên (đã bỏ dấu, lower)."""
    col_map = {
        "med_name": None,      # Tên thuốc hoặc Tên sản phẩm
        "active": None,        # Hoạt chất
        "strength": None,      # Hàm lượng/Nồng độ
        "dosage_form": None,   # Dạng bào chế
        "route": None,         # Đường dùng
        "quantity": None,      # Số lượng
        "unit": None,          # Đơn vị tính
        "plan_price": None,    # Giá kế hoạch
        "award_price": None,   # Đơn giá trúng thầu
        "total_amount": None,  # Thành tiền
        "winner": None,        # Nhà thầu trúng thầu
        "customer": None,      # Khách hàng phụ trách
        "region": None,        # Miền
        "zone": None,          # Vùng
        "province": None,      # Tỉnh
        "hospital": None,      # Bệnh viện/SYT
        "drug_group": None     # Nhóm thuốc (nếu có)
    }
    for col in df.columns:
        col_norm = remove_accents(str(col)).lower()
        if "hoat chat" in col_norm or "thanh phan" in col_norm:
            col_map["active"] = col
        elif "ham luong" in col_norm or "nong do" in col_norm:
            col_map["strength"] = col
        elif "dang bao che" in col_norm:
            col_map["dosage_form"] = col
        elif "duong dung" in col_norm:
            col_map["route"] = col
        elif col_norm.startswith("ten thuoc") or "ten thuoc" in col_norm:
            if col_map["med_name"] is None:
                col_map["med_name"] = col
        elif "ten san pham" in col_norm:
            if col_map["med_name"] is None:
                col_map["med_name"] = col
        elif col_norm.startswith("so luong") or col_norm == "so luong":
            col_map["quantity"] = col
        elif "don vi" in col_norm and "tinh" in col_norm:
            col_map["unit"] = col
        elif "gia ke hoach" in col_norm or "gia moi thau" in col_norm or "gia du kien" in col_norm:
            col_map["plan_price"] = col
        elif "don gia trung thau" in col_norm or "gia trung thau" in col_norm:
            col_map["award_price"] = col
        elif "thanh tien" in col_norm:
            col_map["total_amount"] = col
        elif "nha thau" in col_norm and "trung thau" in col_norm:
            col_map["winner"] = col
        elif "khach hang" in col_norm and "phu trach" in col_norm:
            col_map["customer"] = col
        elif col_norm == "mien":
            col_map["region"] = col
        elif col_norm == "vung":
            col_map["zone"] = col
        elif col_norm == "tinh":
            col_map["province"] = col
        elif "benh vien" in col_norm or col_norm.startswith("so y te"):
            col_map["hospital"] = col
        elif "nhom thuoc" in col_norm:
            col_map["drug_group"] = col
    return col_map

def format_number(num):
    """Định dạng số liệu (tiền tệ hoặc số lượng) để hiển thị."""
    try:
        value = float(num)
    except:
        # nếu không phải số thì trả về chuỗi gốc
        return str(num)
    if value >= 1e9:
        return f"{value/1e9:.2f} tỷ"
    elif value >= 1e6:
        return f"{value/1e6:.1f} triệu"
    elif value >= 1000:
        return f"{int(value):,}"
    else:
        return str(int(value) if value.is_integer() else round(value, 2))

# Giao diện sidebar tải file
st.sidebar.header("📁 Chọn File Dữ Liệu")
tender_file = st.sidebar.file_uploader("Danh mục mời thầu (Excel)", type=["xlsx", "xls", "csv"])
company_file = st.sidebar.file_uploader("Danh mục sản phẩm công ty (Excel/CSV)", type=["xlsx", "xls", "csv"])
assign_file = st.sidebar.file_uploader("File phân công KH phụ trách (tùy chọn)", type=["xlsx", "xls", "csv"])
awarded_file = st.sidebar.file_uploader("Danh mục trúng thầu (Excel)", type=["xlsx", "xls", "csv"])

# Tạo các tab
tab1, tab2, tab3 = st.tabs(["🔎 Lọc Danh Mục Thầu", "📊 Phân Tích Danh Mục Thầu", "🏆 Phân Tích Danh Mục Trúng Thầu"])

# Tab 1: Lọc danh mục mời thầu
with tab1:
    st.subheader("🔎 Lọc Danh Mục Thầu")
    st.write("Hãy tải lên **danh mục mời thầu** và **danh mục sản phẩm công ty** ở thanh bên để bắt đầu lọc.")
    if tender_file is not None and company_file is not None:
        # Đọc dữ liệu
        df_tender, tender_header = load_excel(tender_file)
        df_company, comp_header = load_excel(company_file)
        tender_cols = identify_columns(df_tender)
        comp_cols = identify_columns(df_company)
        # Kiểm tra cột bắt buộc
        if not tender_cols["active"] or not tender_cols["strength"]:
            st.error("Không tìm thấy cột *hoạt chất* và *hàm lượng* trong file mời thầu. Vui lòng kiểm tra lại file.")
            st.stop()
        if not comp_cols["active"] or not comp_cols["strength"]:
            st.error("Không tìm thấy cột *hoạt chất* và *hàm lượng* trong file sản phẩm công ty. Vui lòng kiểm tra lại file.")
            st.stop()
        # Xác định tên cột sản phẩm trong danh mục công ty (nếu có)
        comp_active_col = comp_cols["active"]
        comp_strength_col = comp_cols["strength"]
        comp_product_col = comp_cols["med_name"] if comp_cols["med_name"] else comp_active_col
        # Tạo từ điển key -> tên sản phẩm công ty
        company_dict = {}
        for _, row in df_company.iterrows():
            active_str = str(row[comp_active_col]) if pd.notna(row[comp_active_col]) else ""
            strength_str = str(row[comp_strength_col]) if pd.notna(row[comp_strength_col]) else ""
            key = normalize_text(active_str + strength_str)
            if key and key not in company_dict:
                prod_name = str(row[comp_product_col]) if pd.notna(row[comp_product_col]) else (active_str + " " + strength_str)
                company_dict[key] = prod_name
        # Lọc các mục thầu khớp với danh mục công ty
        output_df = df_tender.copy()
        match_col = "Sản phẩm_Công ty khớp"
        output_df[match_col] = ""
        for idx, row in df_tender.iterrows():
            active_str = str(row[tender_cols["active"]]) if pd.notna(row[tender_cols["active"]]) else ""
            strength_str = str(row[tender_cols["strength"]]) if pd.notna(row[tender_cols["strength"]]) else ""
            tender_key = normalize_text(active_str + strength_str)
            if tender_key in company_dict:
                output_df.at[idx, match_col] = company_dict[tender_key]
        # Hiển thị kết quả lọc (các dòng có khớp)
        df_matched = output_df[output_df[match_col] != ""]
        count = df_matched.shape[0]
        st.write(f"**Kết quả:** Có {count} mặt hàng trong danh mục mời thầu khớp với danh mục sản phẩm của công ty.")
        st.dataframe(df_matched, height=400)
        # Nút tải file Excel kết quả
        try:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name="Ket_qua_loc")
            st.download_button(label="💾 Tải kết quả lọc (Excel)", 
                               data=output.getvalue(), 
                               file_name="Ketqua_loc_danhmucthau.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Xuất Excel thất bại: {e}")
    else:
        st.info("🛈 Vui lòng tải **cả hai file** (danh mục mời thầu và danh mục sản phẩm công ty) để thực hiện lọc.")

# Tab 2: Phân tích danh mục thầu
with tab2:
    st.subheader("📊 Phân Tích Danh Mục Thầu")
    if tender_file is not None:
        # Đảm bảo đã có df_tender
        if 'df_tender' not in locals():
            df_tender, tender_header = load_excel(tender_file)
        tender_cols = identify_columns(df_tender)
        # Xác định các cột liên quan
        active_col = tender_cols["active"]
        strength_col = tender_cols["strength"]
        route_col = tender_cols["route"]
        quantity_col = tender_cols["quantity"]
        plan_price_col = tender_cols["plan_price"]
        total_col = tender_cols["total_amount"]
        # Tính cột tổng giá trị kế hoạch nếu cần
        df_analysis = df_tender.copy()
        if total_col and total_col in df_analysis.columns:
            df_analysis[total_col] = pd.to_numeric(df_analysis[total_col], errors='coerce')
        else:
            if plan_price_col and quantity_col:
                df_analysis[plan_price_col] = pd.to_numeric(df_analysis[plan_price_col], errors='coerce')
                df_analysis[quantity_col] = pd.to_numeric(df_analysis[quantity_col], errors='coerce')
                df_analysis["TongGiaKeHoach"] = df_analysis[plan_price_col] * df_analysis[quantity_col]
                total_col = "TongGiaKeHoach"
            else:
                total_col = None
        if quantity_col:
            df_analysis[quantity_col] = pd.to_numeric(df_analysis[quantity_col], errors='coerce')
        # Biểu đồ Top 10 hoạt chất theo đường dùng (tiêm/uống)
        if active_col and route_col and total_col:
            # Tách dữ liệu theo đường dùng
            df_inj = df_analysis[df_analysis[route_col].astype(str).str.contains("tiêm", case=False, na=False)]
            df_oral = df_analysis[df_analysis[route_col].astype(str).str.contains("uống", case=False, na=False)]
            # Nhóm theo hoạt chất (lowercase để nhóm chính xác)
            if not df_inj.empty:
                df_inj_group = df_inj.copy()
                df_inj_group['active_lower'] = df_inj_group[active_col].astype(str).str.lower()
                df_inj_group = df_inj_group.groupby('active_lower').agg({total_col: 'sum', quantity_col: 'sum', active_col: 'first'}).reset_index(drop=True)
                df_inj_top = df_inj_group.sort_values(total_col, ascending=False).head(10)
                # Nhãn hoạt chất (viết hoa chữ cái đầu cho đẹp)
                df_inj_top[active_col] = df_inj_top[active_col].str.title()
                # Tạo cột text hiển thị cả giá trị và số lượng
                df_inj_top["text"] = df_inj_top.apply(lambda r: f"{format_number(r[total_col])} ({format_number(r[quantity_col])})", axis=1)
                fig_inj = px.bar(df_inj_top, x=active_col, y=total_col, text="text", title="Top 10 hoạt chất (đường tiêm)")
                fig_inj.update_traces(textposition='outside')
                fig_inj.update_yaxes(title="Tổng giá trị kế hoạch (VND)")
                fig_inj.update_xaxes(title="Hoạt chất (đường tiêm)")
                st.plotly_chart(fig_inj, use_container_width=True)
            if not df_oral.empty:
                df_oral_group = df_oral.copy()
                df_oral_group['active_lower'] = df_oral_group[active_col].astype(str).str.lower()
                df_oral_group = df_oral_group.groupby('active_lower').agg({total_col: 'sum', quantity_col: 'sum', active_col: 'first'}).reset_index(drop=True)
                df_oral_top = df_oral_group.sort_values(total_col, ascending=False).head(10)
                df_oral_top[active_col] = df_oral_top[active_col].str.title()
                df_oral_top["text"] = df_oral_top.apply(lambda r: f"{format_number(r[total_col])} ({format_number(r[quantity_col])})", axis=1)
                fig_oral = px.bar(df_oral_top, x=active_col, y=total_col, text="text", title="Top 10 hoạt chất (đường uống)")
                fig_oral.update_traces(textposition='outside')
                fig_oral.update_yaxes(title="Tổng giá trị kế hoạch (VND)")
                fig_oral.update_xaxes(title="Hoạt chất (đường uống)")
                st.plotly_chart(fig_oral, use_container_width=True)
        else:
            st.warning("Không đủ dữ liệu để thống kê Top 10 hoạt chất (thiếu cột hoạt chất, đường dùng hoặc giá trị).")
        # Biểu đồ phân tích theo khách hàng phụ trách (nếu có dữ liệu)
        if assign_file is not None or tender_cols["customer"]:
            if assign_file is not None:
                df_assign, assign_header = load_excel(assign_file)
            else:
                df_assign = df_analysis  # trường hợp cột KH phụ trách nằm ngay trong df_tender
            assign_cols = identify_columns(df_assign)
            customer_col = assign_cols["customer"]
            product_col = assign_cols["med_name"] or assign_cols["active"]
            if customer_col:
                # Nếu file phân công có cột đơn vị (BV/SYT), cố gắng lọc theo đơn vị của gói thầu hiện tại
                if assign_cols["hospital"] and tender_cols["hospital"]:
                    hosp_name = str(df_analysis[tender_cols["hospital"]].iloc[0]) if tender_cols["hospital"] else ""
                    if hosp_name:
                        df_assign = df_assign[df_assign[assign_cols["hospital"]].astype(str).str.contains(hosp_name, case=False, na=False)]
                # Gộp thông tin khách hàng phụ trách vào danh mục thầu theo tên sản phẩm
                df_merge = df_analysis.copy()
                if product_col in df_merge.columns and product_col in df_assign.columns:
                    # nối theo tên sản phẩm/thuốc
                    df_merge = df_merge.merge(df_assign[[product_col, customer_col]], on=product_col, how='left')
                elif tender_cols["active"] and assign_cols["active"] and assign_cols["strength"]:
                    # nếu không có cột tên sản phẩm chung, thử nối theo hoạt chất+hàm lượng (không chắc nhưng thử)
                    df_assign["key"] = df_assign[assign_cols["active"]].astype(str).str.lower() + df_assign[assign_cols["strength"]].astype(str).str.lower()
                    df_merge["key"] = df_merge[tender_cols["active"]].astype(str).str.lower() + df_merge[tender_cols["strength"]].astype(str).str.lower()
                    df_merge = df_merge.merge(df_assign[["key", customer_col]], on="key", how="left")
                # Nhóm theo khách hàng và tính tổng
                df_merge[total_col] = pd.to_numeric(df_merge[total_col], errors='coerce')
                if quantity_col:
                    df_merge[quantity_col] = pd.to_numeric(df_merge[quantity_col], errors='coerce')
                df_by_cust = df_merge.groupby(customer_col).agg({total_col: 'sum', quantity_col: 'sum'}).reset_index()
                df_by_cust = df_by_cust.dropna(subset=[customer_col])
                if df_by_cust.empty:
                    st.info("Không có dữ liệu phân công phù hợp để phân tích theo khách hàng.")
                else:
                    df_by_cust = df_by_cust.sort_values(total_col, ascending=False)
                    df_by_cust["Tổng trị giá (VND)"] = df_by_cust[total_col]
                    df_by_cust["Tổng số lượng"] = df_by_cust[quantity_col]
                    df_cust_melt = df_by_cust.melt(id_vars=customer_col, value_vars=["Tổng trị giá (VND)", "Tổng số lượng"], 
                                                   var_name="Chỉ tiêu", value_name="Giá trị")
                    fig_cust = px.bar(df_cust_melt, x=customer_col, y="Giá trị", color="Chỉ tiêu", barmode="group",
                                      title="Phân tích theo khách hàng phụ trách")
                    fig_cust.update_traces(text=df_cust_melt["Giá trị"].apply(format_number), textposition='outside')
                    fig_cust.update_xaxes(title="Khách hàng phụ trách")
                    fig_cust.update_yaxes(title=None)
                    st.plotly_chart(fig_cust, use_container_width=True)
            else:
                st.warning("File phân công không có thông tin khách hàng phụ trách phù hợp.")
        else:
            st.info("🛈 Có thể tải file phân công khách hàng (nếu có) để xem thống kê theo người phụ trách.")
        # Tra cứu hoạt chất
        st.markdown("---")
        st.subheader("🔍 Tra cứu hoạt chất trong danh mục")
        query = st.text_input("Nhập tên hoạt chất hoặc từ khóa:")
        if query:
            if active_col:
                result_df = df_analysis[df_analysis[active_col].astype(str).str.lower().str.contains(query.strip().lower())]
                if result_df.empty:
                    st.write("Không tìm thấy hoạt chất phù hợp.")
                else:
                    cols_to_show = []
                    for col_key in ["active", "strength", "dosage_form", "route", "drug_group", "plan_price", "total_amount", "unit"]:
                        if tender_cols[col_key]:
                            cols_to_show.append(tender_cols[col_key])
                    cols_to_show = list(dict.fromkeys(cols_to_show))  # loại bỏ trùng lặp nếu có
                    st.dataframe(result_df[cols_to_show].reset_index(drop=True))
            else:
                st.warning("Không xác định được cột hoạt chất để tra cứu trong danh mục.")
    else:
        st.info("Vui lòng tải file danh mục mời thầu ở thanh bên để xem phân tích.")

# Tab 3: Phân tích danh mục trúng thầu
with tab3:
    st.subheader("🏆 Phân Tích Danh Mục Trúng Thầu")
    if awarded_file is not None:
        df_award, award_header = load_excel(awarded_file)
        award_cols = identify_columns(df_award)
        active_col = award_cols["active"]
        if not active_col:
            st.error("Không tìm thấy cột hoạt chất trong file trúng thầu (tiêu đề có thể khác, vui lòng kiểm tra).")
            st.stop()
        quantity_col = award_cols["quantity"]
        award_price_col = award_cols["award_price"]
        total_col = award_cols["total_amount"]
        # Tính thành tiền nếu chưa có
        if total_col:
            df_award[total_col] = pd.to_numeric(df_award[total_col], errors='coerce')
        elif quantity_col and award_price_col:
            df_award[award_price_col] = pd.to_numeric(df_award[award_price_col], errors='coerce')
            df_award[quantity_col] = pd.to_numeric(df_award[quantity_col], errors='coerce')
            df_award["ThanhTien_TinhToan"] = df_award[award_price_col] * df_award[quantity_col]
            total_col = "ThanhTien_TinhToan"
        # Bộ lọc vùng miền
        df_filtered = df_award.copy()
        region_col = award_cols["region"]; zone_col = award_cols["zone"]
        province_col = award_cols["province"]; hospital_col = award_cols["hospital"]
        if region_col:
            regions = ["Tất cả"] + sorted(df_award[region_col].dropna().unique().tolist())
            sel_region = st.selectbox("Chọn Miền:", regions)
            if sel_region and sel_region != "Tất cả":
                df_filtered = df_filtered[df_filtered[region_col] == sel_region]
        if zone_col:
            zones = ["Tất cả"] + sorted(df_filtered[zone_col].dropna().unique().tolist())
            sel_zone = st.selectbox("Chọn Vùng:", zones)
            if sel_zone and sel_zone != "Tất cả":
                df_filtered = df_filtered[df_filtered[zone_col] == sel_zone]
        if province_col:
            provinces = ["Tất cả"] + sorted(df_filtered[province_col].dropna().unique().tolist())
            sel_prov = st.selectbox("Chọn Tỉnh:", provinces)
            if sel_prov and sel_prov != "Tất cả":
                df_filtered = df_filtered[df_filtered[province_col] == sel_prov]
        if hospital_col:
            hospitals = ["Tất cả"] + sorted(df_filtered[hospital_col].dropna().unique().tolist())
            sel_hosp = st.selectbox("Chọn Bệnh viện/SYT:", hospitals)
            if sel_hosp and sel_hosp != "Tất cả":
                df_filtered = df_filtered[df_filtered[hospital_col] == sel_hosp]
        st.write(f"**Số mặt hàng:** {df_filtered.shape[0]}")
        st.dataframe(df_filtered.head(50))
        # Biểu đồ top 10 hoạt chất theo giá trị trúng thầu (trong phạm vi đã lọc)
        if total_col:
            df_group = df_filtered.copy()
            df_group['active_lower'] = df_group[active_col].astype(str).str.lower()
            df_group = df_group.groupby('active_lower').agg({total_col: 'sum', active_col: 'first'}).reset_index(drop=True)
            df_top10 = df_group.sort_values(total_col, ascending=False).head(10)
            df_top10[active_col] = df_top10[active_col].str.title()
            fig_top10 = px.bar(df_top10, x=active_col, y=total_col, 
                                text=df_top10[total_col].apply(format_number),
                                title="Top 10 hoạt chất trúng thầu (giá trị)")
            fig_top10.update_traces(textposition='outside')
            fig_top10.update_xaxes(title="Hoạt chất")
            fig_top10.update_yaxes(title="Tổng trị giá trúng thầu (VND)")
            st.plotly_chart(fig_top10, use_container_width=True)
        else:
            st.info("Không có dữ liệu thành tiền để thống kê.")
    else:
        st.info("🛈 Vui lòng tải file danh mục trúng thầu để xem phân tích.")
