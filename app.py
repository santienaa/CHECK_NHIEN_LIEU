import streamlit as st
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import openpyxl
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="Phân Bổ Nhiên Liệu", layout="wide")

FILE_EXCEL = "DATA.xlsx"

# --- CÁC HÀM XỬ LÝ LOGIC ---
def parse_date_input(date_str):
    """Hàm hỗ trợ ép kiểu chuỗi gõ tay thành datetime"""
    date_str = str(date_str).strip()
    formats = ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    return None

@st.cache_data(ttl=3600)
def load_data():
    if not os.path.exists(FILE_EXCEL):
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_xe = pd.read_excel(FILE_EXCEL, sheet_name="THONG_TIN_XE")
        df_gia = pd.read_excel(FILE_EXCEL, sheet_name="CHI_TIET_GIA_NHIEN_LIEU")
        df_gia["Ngày"] = pd.to_datetime(df_gia["Ngày"], dayfirst=True)
        return df_xe, df_gia
    except Exception as e:
        st.error(f"Lỗi đọc file Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()

# Giao diện Streamlit
st.title("⛽ Hệ Thống Phân Bổ Nhiên Liệu & Cập Nhật Giá Tự Động")

df_xe, df_gia = load_data()

if df_xe.empty:
    st.warning("⚠️ Không tìm thấy dữ liệu xe trong file `DATA.xlsx`. Vui lòng kiểm tra lại file.")
else:
    col1, col2, col3 = st.columns(3)
    
    # Cho phép gõ tay dạng DD/MM/YYYY hoặc chọn lịch
    default_from = (datetime.now() - timedelta(days=30)).strftime("%d/%m/%Y")
    default_to = datetime.now().strftime("%d/%m/%Y")

    with col1:
        str_from = st.text_input("Từ ngày (gõ dd/mm/yyyy)", value=default_from, help="Ví dụ: 11/10/2026")
    with col2:
        str_to = st.text_input("Đến ngày (gõ dd/mm/yyyy)", value=default_to, help="Ví dụ: 10/11/2026")
    with col3:
        danh_sach_xe = df_xe["BIEN SO XE"].dropna().tolist()
        bien_so = st.selectbox("Chọn Biển Số Xe", danh_sach_xe)

    # Chuyển đổi ngày gõ tay sang kiểu Date
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

        # Xác định loại nhiên liệu
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

        # Lọc dữ liệu giá theo khoảng thời gian (Sử dụng đúng tên cột 'Giá nhiên liệu')
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

        # Hiển thị bảng tính toán
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
        
        # Bảng cho phép chỉnh sửa số lít nhập vào
        edited_df = st.data_editor(df_table, num_rows="fixed", use_container_width=True)

        # Tính toán kết quả tổng hợp
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
