import streamlit as st
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import openpyxl
import re
from datetime import datetime, timedelta
import easyocr

st.set_page_config(page_title="Phân Bổ Nhiên Liệu", layout="wide")

FILE_EXCEL = "DATA.xlsx"
PETROLIMEX_URL = "https://www.petrolimex.com.vn/ndi/thong-cao-bao-chi.html"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36"
}

# --- CÁC HÀM XỬ LÝ LOGIC ---
def standardize_item_name(name):
    name_lower = name.lower()
    if "95" in name_lower and ("v" in name_lower or "5" in name_lower or "ei" in name_lower or "e10" in name_lower):
        if "3" not in name_lower and "mức 3" not in name_lower:
            return "E10 RON 95-V"
    if "0,001" in name or "0.001" in name or "001" in name:
        return "DO 001"
    if "0,05" in name or "0.05" in name or "05" in name or "điêzen" in name_lower:
        if "0,001" not in name and "0.001" not in name and "001" not in name:
            return "DO 05"
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
    
    with col1:
        tu_ngay = st.date_input("Từ ngày", datetime.now() - timedelta(days=30))
    with col2:
        den_ngay = st.date_input("Đến ngày", datetime.now())
    with col3:
        danh_sach_xe = df_xe["BIEN SO XE"].dropna().tolist()
        bien_so = st.selectbox("Chọn Biển Số Xe", danh_sach_xe)

    if bien_so:
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

        # Lọc dữ liệu giá theo khoảng thời gian
        data_in_range = df_gia[
            (df_gia["Mặt hàng"] == fuel) &
            (df_gia["Ngày"] >= tu_ts) &
            (df_gia["Ngày"] <= den_ts)
        ][["Ngày", "Giá Vùng 1"]]

        df_b = df_gia[(df_gia["Mặt hàng"] == fuel) & (df_gia["Ngày"] < tu_ts)]
        data_before = pd.DataFrame(columns=["Ngày", "Giá Vùng 1"])
        if not df_b.empty:
            max_d = df_b["Ngày"].max()
            data_before = df_b[df_b["Ngày"] == max_d][["Ngày", "Giá Vùng 1"]]

        combined = pd.concat([data_before, data_in_range]).drop_duplicates().sort_values("Ngày")

        # Hiển thị bảng chỉnh sửa phân bổ nhiên liệu trực tiếp trên web
        st.subheader("Bảng tính toán phân bổ chi tiết")
        
        # Tạo bảng hiển thị với cột Phân bổ mặc định là 0 hoặc người dùng nhập
        table_data = []
        for _, r in combined.iterrows():
            table_data.append({
                "Nhiên liệu phân bổ (lít)": 0.0,
                "Số KM quy đổi": 0.0,
                "Ngày khớp": r["Ngày"].strftime("%d/%m/%Y"),
                "Giá nhiên liệu (VNĐ)": r["Giá Vùng 1"]
            })

        df_table = pd.DataFrame(table_data)
        
        # Sử dụng data_editor để người dùng có thể nhập liệu trực tiếp trên web giống như Excel/Treeview
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