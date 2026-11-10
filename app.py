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

# --- HÀM TỰ ĐỘNG CẬP NHẬT GIÁ TỪ PETROLIMEX ---
def update_petrolimex_prices():
    try:
        sheet_name = "CHI_TIET_GIA_NHIEN_LIEU"
        last_recorded_date = None
        
        if os.path.exists(FILE_EXCEL):
            wb_check = openpyxl.load_workbook(FILE_EXCEL, data_only=True)
            if sheet_name in wb_check.sheetnames:
                ws_check = wb_check[sheet_name]
                date_col_idx = 3
                for col in range(1, ws_check.max_column + 1):
                    if str(ws_check.cell(row=1, column=col).value).strip() == "Ngày":
                        date_col_idx = col
                        break
                for r in range(ws_check.max_row, 0, -1):
                    val = ws_check.cell(row=r, column=date_col_idx).value
                    if val is not None and str(val).strip() != "":
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

        response = requests.get(PETROLIMEX_URL, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        articles_to_process = []
        for a in soup.find_all('a', href=True):
            href = a['href']
            title_text = a.get_text(strip=True)
            if "dieu-chinh-gia-xang-dau" in href or "điều chỉnh giá" in title_text.lower():
                article_link = urljoin("https://www.petrolimex.com.vn", href)
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
            return

        reader = easyocr.Reader(['vi', 'en'], gpu=False)
        all_new_rows = []

        for art in articles_to_process:
            res_article = requests.get(art["link"], headers=HEADERS, timeout=10)
            if res_article.status_code != 200:
                continue
            
            soup_article = BeautifulSoup(res_article.text, 'html.parser')
            page_title = soup_article.find('title').text if soup_article.find('title') else art["title"]
            
            date_match = re.search(r'ngày\s+(\d{1,2}[./]\d{1,2}[./]\d{4})', page_title)
            adjustment_date = datetime.strptime(date_match.group(1).replace('/', '.'), "%d.%m.%Y") if date_match else datetime.now()
            adjustment_date = adjustment_date + timedelta(days=1)

            img_url = None
            for img in soup_article.find_all('img'):
                src = img.get('src', '')
                if "files.petrolimex.com.vn" in src and "s256" not in src and any(ext in src for ext in ["jpg", "png", "webp"]):
                    img_url = urljoin("https://www.petrolimex.com.vn", src)
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
                    all_new_rows.append({
                        "Mặt hàng": standard_name,
                        "Giá nhiên liệu": price_vung_1,
                        "Ngày": adjustment_date
                    })

        if not all_new_rows:
            return

        df_new = pd.DataFrame(all_new_rows).sort_values(by="Ngày", ascending=True)

        wb = openpyxl.load_workbook(FILE_EXCEL)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            last_real_row = 1
            for r in range(ws.max_row, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for c in range(1, ws.max_column + 1)):
                    last_real_row = r
                    break
            start_row = last_real_row + 1
        else:
            ws = wb.create_sheet(sheet_name)
            ws.cell(row=1, column=1, value="Mặt hàng")
            ws.cell(row=1, column=2, value="Giá nhiên liệu")
            ws.cell(row=1, column=3, value="Ngày")
            start_row = 2

        current_row = start_row
        for _, row_data in df_new.iterrows():
            ws.cell(row=current_row, column=1, value=row_data["Mặt hàng"])
            cell_price = ws.cell(row=current_row, column=2, value=row_data["Giá nhiên liệu"])
            cell_price.number_format = '#,##0'
            cell_date = ws.cell(row=current_row, column=3, value=row_data["Ngày"])
            cell_date.number_format = 'DD/MM/YYYY'
            current_row += 1

        wb.save(FILE_EXCEL)
    except Exception as e:
        print(f"Lỗi cào dữ liệu Petrolimex: {e}")

# --- ĐỌC DỮ LIỆU EXCEL ---
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

# Tự động chạy cập nhật giá khi app khởi chạy
with st.spinner("Đang đồng bộ và cập nhật giá nhiên liệu từ Petrolimex..."):
    update_petrolimex_prices()

df_xe, df_gia = load_data()

def parse_date_input(date_str):
    date_str = str(date_str).strip()
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d.%m.%Y"]:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            pass
    return None

# --- GIAO DIỆN STREAMLIT ---
st.title("⛽ Hệ Thống Phân Bổ Nhiên Liệu & Cập Nhật Giá Tự Động")

if df_xe.empty:
    st.warning("⚠️ Không tìm thấy dữ liệu xe trong file `DATA.xlsx`.")
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
