import streamlit as st
import requests
import pandas as pd
from datetime import datetime
import os
import io
import urllib3
import unicodedata
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ====== CẤU HÌNH EMAIL (đọc từ Streamlit Secrets) ======
SENDER_EMAIL    = st.secrets["SENDER_EMAIL"]
SENDER_APP_PASS = st.secrets["SENDER_APP_PASS"]

MODE_SEPARATE   = "separate"
MODE_MULTISHEET = "multisheet"
MODE_ONESHEET   = "onesheet"

EMAIL_RECIPIENTS = [
    ("beatme", "duylinhvinhphuc@gmail.com"),
    # ... (giữ nguyên danh sách)
]

# ====== Giữ nguyên các hàm ======
# remove_accents(), process_data_for_excel(), _clean_df(),
# save_separate(), save_multisheet(), save_onesheet(),
# fetch_keyword_raw(), send_email_to(), run_all()
# → Copy y chang từ file gốc, KHÔNG đổi gì

# ====== GIAO DIỆN STREAMLIT ======
st.set_page_config(page_title="Dữ liệu muasamcong", page_icon="🔍")
st.title("🔍 Tra cứu dữ liệu Mua Sắm Công")

# Input
keyword_input = st.text_input(
    "Từ khóa hàng hóa (nhiều từ khóa ngăn cách bằng dấu ';')",
    value="iphone; máy chủ; laptop"
)
total_pages = st.number_input("Số trang / từ khóa (50 bản ghi/trang)", 
                               min_value=1, max_value=500, value=10)

save_mode = st.radio("Chế độ lưu file", [
    ("Mỗi từ khóa → 1 file Excel riêng", MODE_SEPARATE),
    ("Gộp – mỗi từ khóa 1 sheet", MODE_MULTISHEET),
    ("Gộp – tất cả vào 1 sheet (có cột 'tu_khoa')", MODE_ONESHEET),
], format_func=lambda x: x[0])
save_mode_val = save_mode[1]

# Chọn người nhận
recip_options = [f"{name.strip()} <{email}>" for name, email in EMAIL_RECIPIENTS]
selected_recip = st.multiselect("📧 Gửi email tới:", recip_options)
selected_recipients = [
    EMAIL_RECIPIENTS[recip_options.index(s)] for s in selected_recip
]

# Nút chạy
if st.button("🚀 LẤY DỮ LIỆU", type="primary"):
    keywords = [k.strip() for k in keyword_input.split(';') if k.strip()]
    if not keywords:
        st.error("Vui lòng nhập ít nhất 1 từ khóa!")
    else:
        log_area = st.empty()
        log_lines = []

        def log(msg):
            log_lines.append(msg)
            log_area.text("\n".join(log_lines))

        results_summary, all_kw_data = [], []

        for kw in keywords:
            log(f"\n🔍 Tìm kiếm: \"{kw}\"")
            data = fetch_keyword_raw(kw, total_pages, log)
            if data:
                log(f"  ✅ \"{kw}\": {len(data)} bản ghi")
                all_kw_data.append((kw, data))
                results_summary.append({"keyword": kw, "count": len(data)})
            else:
                log(f"  ⚠️ \"{kw}\": không có kết quả")
                results_summary.append({"keyword": kw, "count": 0})

        if all_kw_data:
            # Lưu file vào bộ nhớ (BytesIO) để download
            buffer = io.BytesIO()
            if save_mode_val == MODE_ONESHEET:
                frames = [_clean_df(data, kw) for kw, data in all_kw_data]
                df = pd.concat(frames, ignore_index=True)
                with pd.ExcelWriter(buffer, engine='openpyxl') as w:
                    df.to_excel(w, sheet_name='msc_data', index=False)
            # (tương tự cho SEPARATE và MULTISHEET)
            
            buffer.seek(0)
            fname = f"mscdata_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            
            st.download_button(
                label="📥 Tải file Excel",
                data=buffer,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Gửi email
            if selected_recipients:
                # Cần lưu file tạm để đính kèm email
                tmp_path = f"/tmp/{fname}"
                with open(tmp_path, 'wb') as f:
                    f.write(buffer.getvalue())
                for name, email in selected_recipients:
                    log(f"\n📧 Gửi tới {name}...")
                    ok = send_email_to(name, email, results_summary, [tmp_path], save_mode_val)
                    log("  ✅ Thành công!" if ok else "  ❌ Thất bại!")

        st.success("✅ Hoàn thành!")
