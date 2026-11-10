import streamlit as st
import pandas as pd
import os
import openpyxl  # <-- Đã bổ sung khai báo thư viện openpyxl
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Hệ Thống Phân Bổ Nhiên Liệu", layout="wide")

# --- GIAO DIỆN THANH BÊN (SIDEBAR) ĐỂ UPLOAD FILE ---
st.sidebar.header("📁 Quản lý Dữ liệu")
uploaded_file = st.sidebar.file_uploader("Tải lên file DATA.xlsx mới nhất", type=["xlsx"])

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    excel_source = io.BytesIO(file_bytes)
    st.sidebar.success("✅ Đã sử dụng file Excel bạn vừa tải lên!")
else:
    if os.path.exists("DATA.xlsx"):
        excel_source = "DATA.xlsx"
        with open("DATA.xlsx", "rb") as f:
            file_bytes = f.read()
        st.sidebar.info("ℹ️ Đang dùng file `DATA.xlsx` sẵn có trên hệ thống.")
    else:
        excel_source = None
        file_bytes = None

# --- ĐỌC DỮ LIỆU EXCEL ---
@st.cache_data(ttl=3600)
def load_data(file_src):
    if file_src is None:
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_xe = pd.read_excel(file_src, sheet_name="THONG_TIN_XE")
        df_gia = pd.read_excel(file_src, sheet_name="CHI_TIET_GIA_NHIEN_LIEU")
        
        df_gia.columns = [str(c).strip() for c in df_gia.columns]
        for col in df_gia.columns:
            if "giá" in col.lower():
                df_gia.rename(columns={col: "Giá nhiên liệu"}, inplace=True)
                break

        df_gia["Ngày"] = pd.to_datetime(df_gia["Ngày"], dayfirst=True)
        return df_xe, df_gia
    except Exception as e:
        st.error(f"Lỗi đọc file Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

df_xe, df_gia = load_data(excel_source)

def parse_date_input(date_str):
    date_str = str(date_str).strip()
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    return None

# --- GIAO DIỆN CHÍNH ---
st.title("⛽ Hệ Thống Phân Bổ Nhiên Liệu")

if df_xe.empty or df_gia.empty:
    st.warning("⚠️ Vui lòng tải lên file `DATA.xlsx` hợp lệ ở thanh bên trái (Sidebar) để bắt đầu sử dụng.")
else:
    # --- PHẦN 1: CẬP NHẬT / THÊM MỚI BẢNG GIÁ TRỰC TIẾP TRÊN WEB ---
    with st.expander("➕ Thêm / Cập nhật giá nhiên liệu mới vào hệ thống", expanded=False):
        st.markdown("Bạn có thể chọn mặt hàng từ danh sách có sẵn hoặc tự nhập tên mặt hàng mới, sau đó điền giá và ngày áp dụng.")
        
        danh_sach_mat_hang = df_gia["Mặt hàng"].dropna().unique().tolist()
        
        if "new_price_rows" not in st.session_state:
            st.session_state.new_price_rows = pd.DataFrame(columns=["Mặt hàng", "Giá nhiên liệu", "Ngày"])

        col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 1])
        with col_f1:
            input_mode = st.radio("Cách chọn Mặt hàng:", ["Chọn từ danh sách", "Nhập tay"], horizontal=True)
            if input_mode == "Chọn từ danh sách":
                chon_mh = st.selectbox("Mặt hàng", danh_sach_mat_hang)
            else:
                chon_mh = st.text_input("Nhập tên mặt hàng mới")
        with col_f2:
            nhap_gia = st.text_input("Giá nhiên liệu (VNĐ)", value="0")
        with col_f3:
            nhap_ngay = st.text_input("Ngày áp dụng (dd/mm/yyyy)", value=datetime.now().strftime("%d/%m/%Y"))
        with col_f4:
            st.write("")
            st.write("")
            if st.button("Thêm vào danh sách"):
                try:
                    val_gia = float(nhap_gia.replace(".", "").replace(",", ""))
                except ValueError:
                    val_gia = -1
                
                parsed_d = parse_date_input(nhap_ngay)
                
                if not chon_mh.strip():
                    st.error("❌ Tên mặt hàng không được để trống!")
                elif val_gia < 0:
                    st.error("❌ Giá nhiên liệu phải là số hợp lệ!")
                elif not parsed_d:
                    st.error("❌ Ngày không đúng định dạng dd/mm/yyyy!")
                else:
                    # Sắp xếp đúng thứ tự cột: Mặt hàng, Giá nhiên liệu, Ngày
                    new_row = pd.DataFrame({
                        "Mặt hàng": [chon_mh.strip()], 
                        "Giá nhiên liệu": [val_gia], 
                        "Ngày": [parsed_d.strftime("%d/%m/%Y")]
                    })
                    st.session_state.new_price_rows = pd.concat([st.session_state.new_price_rows, new_row], ignore_index=True)
                    st.success("✅ Đã thêm dòng thành công!")

        if not st.session_state.new_price_rows.empty:
            st.subheader("Danh sách giá mới chờ lưu vào file:")
            st.dataframe(st.session_state.new_price_rows, use_container_width=True)
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("🗑️ Xóa danh sách tạm"):
                    st.session_state.new_price_rows = pd.DataFrame(columns=["Mặt hàng", "Giá nhiên liệu", "Ngày"])
                    st.rerun()
            with col_b2:
                if st.button("📥 Tạo và Tải xuống file DATA.xlsx đã cập nhật"):
                    try:
                        if file_bytes is not None:
                            wb_export = openpyxl.load_workbook(io.BytesIO(file_bytes))
                        else:
                            wb_export = openpyxl.load_workbook("DATA.xlsx")
                        
                        ws_exp = wb_export["CHI_TIET_GIA_NHIEN_LIEU"]
                        
                        for _, r_new in st.session_state.new_price_rows.iterrows():
                            next_r = ws_exp.max_row + 1
                            ws_exp.cell(row=next_r, column=1, value=r_new["Mặt hàng"])
                            c_p = ws_exp.cell(row=next_r, column=2, value=float(r_new["Giá nhiên liệu"]))
                            c_p.number_format = '#,##0'
                            
                            # Chuyển đổi lại định dạng ngày thành kiểu date để Excel hiển thị chuẩn
                            date_obj = datetime.strptime(str(r_new["Ngày"]), "%d/%m/%Y")
                            c_d = ws_exp.cell(row=next_r, column=3, value=date_obj)
                            c_d.number_format = 'DD/MM/YYYY'
                        
                        output_io = io.BytesIO()
                        wb_export.save(output_io)
                        processed_data = output_io.getvalue()
                        
                        st.download_button(
                            label="👉 Bấm vào đây để tải file Excel về máy",
                            data=processed_data,
                            file_name="DATA.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("🎉 File Excel cập nhật đã sẵn sàng để tải xuống!")
                    except Exception as ex:
                        st.error(f"Lỗi khi xuất file: {ex}")

    st.markdown("---")

    # --- PHẦN 2: GIAO DIỆN TRA CỨU VÀ TÍNH TOÁN PHÂN BỔ ---
    col1, col2, col3 = st.columns(3)
    
    default_from = (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y")
    default_to = datetime.now().strftime("%d/%m/%Y")

    with col1:
        str_from = st.text_input("Từ ngày (gõ dd/mm/yyyy)", value=default_from)
    with col2:
        str_to = st.text_input("Đến ngày (gõ dd/mm/yyyy)", value=default_to)
    with col3:
        danh_sach_xe = df_xe["BIEN SO XE"].dropna().tolist()
        bien_so = st.selectbox("Chọn Biển Số Xe", danh_sach_xe)

    tu_ngay = parse_date_input(str_from)
    den_ngay = parse_date_input(str_to)

    if not tu_ngay or not den_ngay:
        st.error("⚠️ Vui lòng nhập ngày đúng định dạng **ngày/tháng/năm** (Ví dụ: 11/10/2026)!")
    elif bien_so:
        matched_xe = df_xe[df_xe["BIEN SO XE"] == bien_so]
        xe = matched_xe.iloc[0]
        dinh_muc = float(xe["DINH MUC NHIEN LIEU (MOI 100 KM)"])
        nl1 = xe["LOAI NHIEN LIEU 1"]
        nl2 = xe["LOAI NHIEN LIEU 2"]

        tu_ts = pd.to_datetime(tu_ngay)
        den_ts = pd.to_datetime(den_ngay)
        
        df_before = df_gia[df_gia["Ngày"] <= tu_ts]
        if df_before.empty:
            df_before = df_gia[df_gia["Ngày"] <= den_ts]

        fuel = nl1
        if not df_before.empty:
            max_date = df_before["Ngày"].max()
            items_on_max_date = df_before[df_before["Ngày"] == max_date]["Mặt hàng"].unique()
            if nl1 in items_on_max_date:
                fuel = nl1
            elif nl2 in items_on_max_date:
                fuel = nl2
            elif len(items_on_max_date) > 0:
                fuel = items_on_max_date[0]

        st.markdown("---")
        m1, m2 = st.columns(2)
        m1.metric("Loại Nhiên Liệu", fuel)
        m2.metric("Định Mức (lít/100km)", f"{dinh_muc:,.2f}")

        data_in_range = df_gia[
            (df_gia["Mặt hàng"] == fuel) &
            (df_gia["Ngày"] >= tu_ts) &
            (df_gia["Ngày"] <= den_ts)
        ][["Ngày", "Giá nhiên liệu"]]

        df_b = df_gia[(df_gia["Mặt hàng"] == fuel) & (df_gia["Ngày"] < tu_ts)]
        data_before = pd.DataFrame(columns=["Ngày", "Giá nhiên liệu"])
        if not df_b.empty:
            max_d = df_b["Ngày"].max()
            data_before = df_b[df_b["Ngày"] == max_d][["Ngày", "Giá nhiên liệu"]]

        combined = pd.concat([data_before, data_in_range]).drop_duplicates().sort_values("Ngày")

        st.subheader("Bảng tính toán phân bổ chi tiết")
        
        table_data = []
        for _, r in combined.iterrows():
            table_data.append({
                "Nhiên liệu phân bổ (lít)": 0.0,
                "Số KM quy đổi": 0.0,
                "Ngày khớp": r["Ngày"].strftime("%d/%m/%Y"),
                "Giá nhiên liệu (VNĐ)": r["Giá nhiên liệu"]
            })

        df_table = pd.DataFrame(table_data)
        edited_df = st.data_editor(df_table, num_rows="fixed", use_container_width=True)

        tong_km = 0
        tong_tien = 0
        
        for index, row in edited_df.iterrows():
            try:
                pb = float(row["Nhiên liệu phân bổ (lít)"])
            except:
                pb = 0.0
            gia = float(row["Giá nhiên liệu (VNĐ)"])
            km = (pb * 100 / dinh_muc) if dinh_muc > 0 else 0
            
            tong_km += km
            tong_tien += (pb * gia)
            edited_df.at[index, "Số KM quy đổi"] = round(km, 2)

        st.markdown("---")
        r1, r2 = st.columns(2)
        r1.metric("Tổng KM quy đổi", f"{tong_km:,.2f} km")
        r2.metric("Tổng tiền nhiên liệu", f"{tong_tien:,.0f} VNĐ".replace(",", "."))
