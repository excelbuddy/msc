import requests
import pandas as pd
from datetime import datetime
import os
import urllib3
import unicodedata
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import streamlit as st
import io

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ====== CẤU HÌNH DANH SÁCH EMAIL ======
EMAIL_RECIPIENTS = [
    ("beatme", "duylinhvinhphuc@gmail.com"),
    ("beatme", "duylinh93@gmail.com"),
    ("Bui Viet Huy", "bvhuy.ho@vietcombank.com.vn"),
    ("Can Quang Minh", "minhcq.ho@vietcombank.com.vn"),
    ("Do Thi Vui", "VUIDT.HO@vietcombank.com.vn"),
    ("Do Thuy Linh", "linhdt1.ho@vietcombank.com.vn"),
    ("Khieu Van Truong", "TRUONGKV.HO@vietcombank.com.vn"),
    ("Luong Trung Hieu", "HIEULT1.HO@vietcombank.com.vn"),
    ("MAI XUAN LICH", "lichmx.ho@vietcombank.com.vn"),
    ("Nguyen Anh Thu", "nathu.ho@vietcombank.com.vn"),
    ("Nguyen Duc Huy", "HUYND.HO@vietcombank.com.vn"),
    ("Nguyen Hoai", "hoain.ho@vietcombank.com.vn"),
    ("Nguyen Kim Nhung", "NHUNGNK.HO@vietcombank.com.vn"),
    ("Nguyen Ngoc Chi Linh", "LINHNNC.HO@vietcombank.com.vn"),
    ("Nguyen Ngoc Khanh", "khanhnn.ho@vietcombank.com.vn"),
    ("Nguyen The Duy", "duynt.ho@vietcombank.com.vn"),
    ("Nguyen Thi Ngoc Bich", "ntnbich.ho2@vietcombank.com.vn"),
    ("Nguyen Thi Ngoc Mai", "MAINTN1.HO@vietcombank.com.vn"),
    ("NGUYEN THU HUONG", "nthuong.ho2@vietcombank.com.vn"),
    ("Nguyen Tu Anh", "NTANH1.HO@vietcombank.com.vn"),
    ("Nguyen Viet Nga", "nganv.ho@vietcombank.com.vn"),
    ("Nguyen Viet Tung", "TUNGNV1.HO@vietcombank.com.vn"),
    ("Pham Kim Ngan", "nganpk.ho@vietcombank.com.vn"),
    ("Pham Thi Thanh Nga", "ngaptt.ho@vietcombank.com.vn"),
    ("Phan Thi Phuong", "phuongpt1.ho@vietcombank.com.vn"),
    ("Phan Thuy Linh", "Linhpt.ho@vietcombank.com.vn"),
    ("Tran Hoang Nga", "ngath.ho1@vietcombank.com.vn"),
    ("Tran Manh Hung", "HUNGTM4.HO@vietcombank.com.vn"),
    ("Tran Quang Anh", "ANHTQ.HO@vietcombank.com.vn"),
    ("Tran Thi Cam Van", "VANTTC.HO@vietcombank.com.vn"),
    ("Tran Thi Huyen Trang", "trangtth.ho2@vietcombank.com.vn"),
    ("Tran Thi Mai Huong", "ttmhuong.ho@vietcombank.com.vn"),
    ("Tran Thi Thanh Hao", "haottt.ho@vietcombank.com.vn"),
    ("Truong Duc Hai", "haitd.ho@vietcombank.com.vn"),
    ("Vu Duy Linh", "LINHVD.HO@vietcombank.com.vn"),
    ("Vu Thi Bich", "BICHVT.HO@vietcombank.com.vn"),
    ("Vu Thu Nga", "ngavt.ho@vietcombank.com.vn"),
    ("Vu Viet Anh", "anhvv1.ho@vietcombank.com.vn"),
]

# Lấy thông tin tài khoản từ Secrets của Streamlit để bảo mật
SENDER_EMAIL = "vuan128903@gmail.com"
try:
    SENDER_APP_PASS = st.secrets["GMAIL_APP_PASS"]
except:
    SENDER_APP_PASS = "jecm hqtm viwa nxyc" # Chạy local tạm thời

MODE_SEPARATE   = "separate"
MODE_MULTISHEET = "multisheet"
MODE_ONESHEET   = "onesheet"

def remove_accents(s):
    nfkd = unicodedata.normalize('NFKD', s)
    s2   = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'[^a-zA-Z0-9_]+', '', s2.replace(' ', '_')).lower()

def process_data_for_excel(response_data):
    if not response_data or 'page' not in response_data or 'content' not in response_data['page']:
        return None
    processed = []
    for item in response_data['page']['content']:
        row = {k: item.get(k, '') for k in item}
        for k in item:
            if isinstance(item[k], list):
                row[k] = '; '.join(str(x) for x in item[k] if x is not None)
        if 'locations' in item and item['locations']:
            row['locations'] = '; '.join(
                f"{loc.get('provName','')} ({loc.get('provCode','')})"
                for loc in item['locations']
            )
        processed.append(row)
    return processed

def _clean_df(data, keyword=None):
    df = pd.DataFrame(data)
    drop_cols = ["id","type","tab","soQuyetDinh","ngayBanHanhQuyetDinh","decisions"]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors='ignore')
    if keyword is not None:
        df.insert(0, 'tu_khoa', keyword)
    return df

def _auto_width(ws):
    for col in ws.columns:
        w, letter = 0, col[0].column_letter
        for cell in col:
            try: w = max(w, len(str(cell.value)))
            except: pass
        ws.column_dimensions[letter].width = min(w + 2, 50)

# ====== LƯU EXCEL TRÊN BỘ NHỚ BIẾN (Để Webapp tải xuống được trực tiếp) ======
def save_separate(data, keyword):
    if not data: return None, 0
    fn = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}_{remove_accents(keyword)}.xlsx"
    try:
        df = _clean_df(data)
        with pd.ExcelWriter(fn, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='msc_data', index=False)
            _auto_width(w.sheets['msc_data'])
        return fn, len(df)
    except Exception as e:
        st.error(f"Lỗi lưu file: {e}")
        return None, 0

def save_multisheet(all_kw_data):
    fn = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}_multisheet.xlsx"
    total = 0
    try:
        with pd.ExcelWriter(fn, engine='openpyxl') as w:
            for kw, data in all_kw_data:
                if not data: continue
                df = _clean_df(data)
                sn = base = remove_accents(kw)[:28] or "sheet"
                i = 2
                while sn in w.sheets: sn = f"{base[:25]}_{i}"; i += 1
                df.to_excel(w, sheet_name=sn, index=False)
                _auto_width(w.sheets[sn])
                total += len(df)
        return fn, total
    except Exception as e:
        st.error(f"Lỗi lưu file: {e}")
        return None, 0

def save_onesheet(all_kw_data):
    fn = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}_onesheet.xlsx"
    try:
        frames = [_clean_df(data, kw) for kw, data in all_kw_data if data]
        if not frames: return None, 0
        df = pd.concat(frames, ignore_index=True)
        with pd.ExcelWriter(fn, engine='openpyxl') as w:
            df.to_excel(w, sheet_name='msc_data', index=False)
            _auto_width(w.sheets['msc_data'])
        return fn, len(df)
    except Exception as e:
        st.error(f"Lỗi lưu file: {e}")
        return None, 0

def send_email_to(recipient_name, recipient_email, results_summary, excel_files):
    try:
        msg = MIMEMultipart()
        msg['From']    = SENDER_EMAIL
        msg['To']      = recipient_email
        msg['Subject'] = f"[Email tự động] Kết quả tra cứu dữ liệu hàng hóa muasamcong - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        lines = [
            f"Xin chào {recipient_name},", 
            "\nLƯU Ý KHI SỬ DỤNG DỮ LIỆU: \n",
            " - Nên dùng các dòng dữ liệu mà cột bidForm có giá trị [DTRR], [CHCT]...",
            " - Dữ liệu chỉ dùng nội bộ trong Ban mua sắm, không gửi ra nơi khác!\n",
            "Kết quả tìm kiếm hàng hóa trên muasamcong:\n"
        ]
        for r in results_summary:
            lines.append(f"  • \"{r['keyword']}\" → {r['count']} bản ghi")
        lines += ["", f"File đính kèm: {len(excel_files)} file.", "", "© Developed by Beatme!"]
        
        msg.attach(MIMEText('\n'.join(lines), 'plain', 'utf-8'))
        
        for fp in excel_files:
            if not os.path.exists(fp): continue
            with open(fp, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(fp)}"')
            msg.attach(part)
            
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(SENDER_EMAIL, SENDER_APP_PASS)
            s.sendmail(SENDER_EMAIL, recipient_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Lỗi gửi email: {e}")
        return False

def fetch_keyword_raw(keyword, total_pages, status_box):
    url = "https://muasamcong.mpi.gov.vn/o/egp-portal-personal-page/services/smart/search_prc"
    headers = {'Accept':'application/json, text/plain, */*', 'Content-Type':'application/json','User-Agent':'Mozilla/5.0'}
    cookies = {'GUEST_LANGUAGE_ID':'vi_VN','COOKIE_SUPPORT':'true'}
    payload = [{"pageSize":50,"pageNumber":0,"query":[{
        "index":"es-smart-pricing","keyWord":keyword,"matchType":"all-1",
        "matchFields":["danh_muc_hang_hoa","ma_hs","xuat_xu","ma_tbmt","ky_ma_hieu","nhan_hieu","hang_san_xuat"],
        "filters":[
            {"fieldName":"type","searchType":"in","fieldValues":["HANG_HOA"]},
            {"fieldName":"tab", "searchType":"in","fieldValues":["HANG_HOA"]}
        ]}]}]
    
    all_data = []
    for page in range(total_pages):
        status_box.write(f"📄 *[{keyword}]* Đang quét trang {page+1}/{total_pages}...")
        payload[0]["pageNumber"] = page
        try:
            r = requests.post(url, headers=headers, cookies=cookies, json=payload, verify=False, timeout=30)
            if r.status_code == 200:
                d = process_data_for_excel(r.json())
                if not d:
                    status_box.write(f"⚠️ *[{keyword}]* Không có dữ liệu tại trang {page+1} $\rightarrow$ Dừng quét.")
                    break
                all_data.extend(d)
                if len(d) < 50:
                    status_box.write(f"✅ *[{keyword}]* Đã tới trang cuối cùng ({len(d)} bản ghi).")
                    break
            else:
                status_box.write(f"❌ *[{keyword}]* Lỗi HTTP {r.status_code}")
        except Exception as e:
            status_box.write(f"❌ *[{keyword}]* Lỗi kết nối: {e}")
    return all_data

# ====== STREAMLIT APP UI ======
st.set_page_config(page_title="Dữ liệu hàng hóa mua sắm công", page_icon="🚀", layout="centered")
st.title("🔍 Tra cứu dữ liệu Mua Sắm Công v5")
st.caption("Developed by Beatme")

# 1. Input Từ khóa
raw_kw = st.text_input("Từ khóa hàng hóa (ngăn cách các từ bằng dấu chấm phẩy ';')", value="iphone; máy chủ; laptop")

# 2. Số trang
total_pages = st.number_input("Số trang tối đa mỗi từ khóa (50 bản ghi/trang):", min_value=1, max_value=500, value=10)

# 3. Chế độ lưu file
save_mode = st.radio(
    "📁 Chế độ xuất Excel:",
    options=[MODE_SEPARATE, MODE_MULTISHEET, MODE_ONESHEET],
    format_func=lambda x: {
        MODE_SEPARATE: "Mỗi từ khóa $\rightarrow$ 1 file Excel riêng",
        MODE_MULTISHEET: "Gộp vào 1 file $\rightarrow$ Mỗi từ khóa = 1 sheet",
        MODE_ONESHEET: "Gộp vào 1 file $\rightarrow$ Tất cả gộp chung 1 sheet (có cột 'tu_khoa')"
    }[x]
)

# 4. Chọn người nhận email
recip_options = [f"{name} <{email}>" for name, email in EMAIL_RECIPIENTS]
selected_recip_strings = st.multiselect("📧 Gửi email báo cáo tới (Có thể để trống nếu chỉ muốn tải file):", options=recip_options)

# Trích xuất ngược lại tuple (name, email) từ multiselect
recipients = []
for r_str in selected_recip_strings:
    for name, email in EMAIL_RECIPIENTS:
        if email in r_str:
            recipients.append((name, email))

# 5. Nút bấm kích hoạt tiến trình
if st.button("🚀 BẮT ĐẦU LẤY DỮ LIỆU", type="primary"):
    keywords = [k.strip() for k in raw_kw.split(';') if k.strip()]
    if not keywords:
        st.error("Vui lòng nhập ít nhất 1 từ khóa!")
    else:
        results_summary, excel_files, all_kw_data = [], [], []
        
        # Tạo khu vực log trạng thái động
        with st.status("🚀 Đang tiến hành lấy dữ liệu...", expanded=True) as status:
            for kw in keywords:
                status.write(f"--- 🔍 Đang tìm kiếm: **{kw}** ---")
                data = fetch_keyword_raw(kw, total_pages, status)
                if data:
                    all_kw_data.append((kw, data))
                    results_summary.append({"keyword": kw, "count": len(data)})
                else:
                    results_summary.append({"keyword": kw, "count": 0})
            
            if not all_kw_data:
                status.update(label="⚠️ Không thu thập được dữ liệu nào!", state="warning")
            else:
                status.write("💾 Đang xuất cấu trúc Excel...")
                if save_mode == MODE_SEPARATE:
                    for kw, data in all_kw_data:
                        fn, cnt = save_separate(data, kw)
                        if fn: excel_files.append(fn)
                elif save_mode == MODE_MULTISHEET:
                    fn, cnt = save_multisheet(all_kw_data)
                    if fn: excel_files.append(fn)
                elif save_mode == MODE_ONESHEET:
                    fn, cnt = save_onesheet(all_kw_data)
                    if fn: excel_files.append(fn)
                
                # Tiến hành gửi Email nếu có chọn người nhận
                if excel_files and recipients:
                    for name, email in recipients:
                        status.write(f"📧 Đang gửi mail tới {name} <{email}>...")
                        if send_email_to(name, email, results_summary, excel_files):
                            status.write(f"✅ Đã gửi thành công cho {name}!")
                        else:
                            status.write(f"❌ Gửi thất bại cho {name}!")
                
                status.update(label="🎉 Hoàn thành xử lý!", state="complete")
        
        # Hiển thị nút Download trực tiếp trên Web sau khi quét xong
        st.success("Tải dữ liệu thành công! Bạn có thể tải trực tiếp các file dưới đây:")
        for fp in excel_files:
            if os.path.exists(fp):
                with open(fp, "rb") as f:
                    st.download_button(
                        label=f"📥 Tải xuống {fp}",
                        data=f,
                        file_name=fp,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )