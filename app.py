import io, smtplib, ssl
from email.message import EmailMessage
import streamlit as st
from openpyxl import load_workbook

XLSX = "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
st.title("📧 엑셀 수정 → Gmail 자동발송")

up      = st.file_uploader("엑셀 파일(.xlsx)", type="xlsx")
cell    = st.text_input("수정할 셀", "B2")
value   = st.text_input("넣을 값", "2026-08-07")
to      = st.text_input("받는 사람 이메일")
subject = st.text_input("제목", "[보고] 월간 실적")
body    = st.text_area("본문", "첨부 확인 부탁드립니다.")

if st.button("수정 후 발송", type="primary", disabled=not (up and to)):
    # 1) 메모리에서 수정
    wb = load_workbook(up)
    wb.active[cell] = value
    buf = io.BytesIO()
    wb.save(buf)

    # 2) 메일 작성
    msg = EmailMessage()
    msg["From"] = st.secrets["GMAIL_USER"]
    msg["To"], msg["Subject"] = to, subject
    msg.set_content(body)
    msg.add_attachment(buf.getvalue(), maintype="application",
                       subtype=XLSX, filename=up.name)

    # 3) 발송
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
            s.login(st.secrets["GMAIL_USER"], st.secrets["GMAIL_APP_PW"])
            s.send_message(msg)
        st.success(f"✅ {to} 로 발송 완료")
        st.download_button("수정본 내려받기", buf.getvalue(), file_name=up.name)
    except Exception as e:
        st.error(f"❌ 실패: {e}")
