import streamlit as st
import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
import openpyxl
import re
from datetime import datetime, timedelta
import easyocr
import io

st.set_page_config(page_title="Phân Bổ Nhiên Liệu & Tự Động Cập Nhật", layout="wide")

PETROLIMEX_URL = "https://www.petrolimex.com.vn/ndi/thong-cao-bao-chi.html"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36"
}

# --- HÀM CHUẨN HÓA MẶT HÀNG ---
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

# --- GIAO DIỆN SIDEBAR UPLOAD FILE ---
st.sidebar.header("📁 Quản lý Dữ liệu")
uploaded_file = st.sidebar.file_uploader("Tải lên file DATA.xlsx của bạn", type=["xlsx"])

# Đọc file gốc từ người dùng upload hoặc file sẵn có
if uploaded_file is not None:
    # Lưu bytes tạm vào session state để xử lý ghi nối dữ liệu
    file_bytes = uploaded_file.read()
    st.sidebar.success("✅ Đã tải file Excel lên thành công!")
else:
    if os.path.exists("DATA.xlsx"):
        with open("DATA.xlsx", "rb") as f:
            file_bytes = f.read()
        st.sidebar.info("ℹ️ Đang dùng file `DATA.xlsx` mặc định trên hệ thống.")
    else:
        file_bytes = None

# --- HÀM CÀO DỮ LIỆU PETROLIMEX VÀ BỔ SUNG VÀO EXCEL ---
def fetch_and_merge_petrolimex(excel_bytes):
    if excel_bytes is None:
        return None, 0, False
    
    try:
        # Đọc workbook từ bytes
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        sheet_name = "CHI_TIET_GIA_NHIEN_LIEU"
        
        if sheet_name not in wb.sheetnames:
            return excel_bytes, 0, False
            
        ws = wb[sheet_name]
        
        # Tìm cột Ngày trong sheet giá
        date_col_idx = 3
        for col in range(1, ws.max_column + 1):
            if str(ws.cell(row=1, column=col).value).strip() == "Ngày":
                date_col_idx = col
                break
                
        # Lấy ngày mới nhất đang có trong file Excel
        last_recorded_date = None
        for r in range(ws.max_row, 1, -1):
            val = ws.cell(row=r, column=date_col_idx).value
            if val is not None:
                if isinstance(val, datetime):
                    last_recorded_date = val.date()
                elif isinstance(val, str):
                    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d.%m.%Y'):
                        try:
                            last_recorded_date = datetime.strptime(val.strip(), fmt).date()
                            break
                        except:
                            pass
                if last_recorded_date:
                    break

        # Cào web Petrolimex
        response = requests.get(PETROLIMEX_URL, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        articles_to_process = []
        for a in soup.find_all('a', href=True):
            href = a['href']
            title_text = a.get_text(strip=True)
            if "dieu-chinh-gia-xang-dau" in href or "điều chỉnh giá" in title_text.lower():
                article_link = urljoin("https://www.petrolimex.com.vn", href) if "http" not in href else href
                date_match = re.search(r'ngày\s+(\d{1,2}[./]\d{1,2}[./]\d{4})', title_text)
                if date_match:
                    raw_date_str = date_match.group(1).replace('/', '.')
                    try:
                        article_date = datetime.strptime(raw_date_str, "%d.%m.%Y").date()
                        actual_effective_date = article_date + timedelta(days=1)
                        if last_recorded_date and actual_effective_date <= last_recorded_date:
                            break 
                    except:
                        pass
                if article_link not in [item['link'] for item in articles_to_process]:
                    articles_to_process.append({"link": article_link, "title": title_text})

        if not articles_to_process:
            return excel_bytes, 0, False

        reader = easyocr.Reader(['vi', 'en'], gpu=False)
        new_items_count = 0

        for art in articles_to_process:
            res_article = requests.get(art["link"], headers=HEADERS, timeout=10)
            if res_article.status_code != 200:
                continue
            
            soup_article = BeautifulSoup(res_article.text, 'html.parser')
            page_title = soup_article.find('title').text if soup_article.find('title') else art["title"]
            
            date_match = re.search(r'ngày\s+(\d{1,2}[./]\d{1,2}[./]\d{4})', page_title)
            adjustment_date = datetime.strptime(date_match.group(1).replace('/', '.'), "%d.%m.%Y") if date_match else datetime.now()
            adjustment_date = adjustment_date + timedelta(days=1)

            # Kiểm tra nếu ngày này đã có trong file thì bỏ qua
            if last_recorded_date and adjustment_date <= last_recorded_date:
                continue

            img_url = None
            for img in soup_article.find_all('img'):
                src = img.get('src', '')
                if "files.petrolimex.com.vn" in src and "s256" not in src and any(ext in src for ext in ["jpg", "png", "webp"]):
                    img_url = src if "http" in src else "https://www.petrolimex.com.vn" + src
                    break

            if not img_url:
                continue

            img_data = requests.get(img_url, headers=HEADERS, timeout=10).content
            temp_img_path = "temp_bang_gia.jpg"
            with open(temp_img_path, 'wb') as f:
                f.write(img_data)

            ocr_results = reader.readtext(temp_img_path, detail=0)
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)

            raw_rows = []
            i = 0
            while i < len(ocr_results) - 3:
                item_name = ocr_results[i]
                p1 = ocr_results[i+2]
                p2 = ocr_results[i+3]
                if "." in p1 and "." in p2 and any(char.isdigit() for char in p1):
                    raw_rows.append([item_name, ocr_results[i+1], p1, p2])
                    i += 4
                else:
                    i += 1

            for row in raw_rows:
                original_name = row[0]
                try:
                    price_vung_1 = int(row[2].replace(".", ""))
                except:
                    price_vung_1 = 0
                
                standard_name = standardize_item_name(original_name)
                if standard_name:
                    # Thêm dòng mới vào cuối sheet
                    next_row = ws.max_row + 1
                    ws.cell(row=next_row, column=1, value=standard_name)
                    cell_price = ws.cell(row=next_row, column=2, value=price_vung_1)
                    cell_price.number_format = '#,##0'
                    cell_date = ws.cell(row=next_row, column=3, value=adjustment_date)
                    cell_date.number_format = 'DD/MM/YYYY'
                    new_items_count += 1

        if new_items_count > 0:
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue(), new_items_count, True
        else:
            return excel_bytes, 0, False

    except Exception as e:
        print(f"Lỗi cào dữ liệu: {e}")
        return excel_bytes, 0, False

# Chạy kiểm tra và bổ sung dữ liệu tự động
with st.spinner("Đang kiểm tra dữ liệu giá mới từ Petrolimex..."):
    final_bytes, added_count, has_new = fetch_and_merge_petrolimex(file_bytes)

if has_new:
    st.sidebar.success(f"✨ Đã tự động cập nhật thêm **{added_count}** bản ghi giá mới từ Petrolimex!")
    st.sidebar.download_button(
        label="📥 Tải xuống file DATA.xlsx mới nhất",
        data=final_bytes,
        file_name="DATA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- ĐỌC DỮ LIỆU ĐỂ HIỂN THỊ ---
@st.cache_data(ttl=3600)
def load_data_from_bytes(b_data):
    if b_data is None:
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_xe = pd.read_excel(io.BytesIO(b_data), sheet_name="THONG_TIN_XE")
        df_gia = pd.read_excel(io.BytesIO(b_data), sheet_name="CHI_TIET_GIA_NHIEN_LIEU")
        
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

df_xe, df_gia = load_data_from_bytes(final_bytes)

def parse_date_input(date_str):
    date_str = str(date_str).strip()
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    return None

# --- GIAO DIỆN CHÍNH ---
st.title("⛽ Hệ Thống Phân Bổ Nhiên Liệu & Cập Nhật Giá Tự Động")

if df_xe.empty:
    st.warning("⚠️ Vui lòng tải lên file `DATA.xlsx` của bạn ở thanh bên trái (Sidebar) để bắt đầu sử dụng.")
else:
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
