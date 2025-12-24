import streamlit as st
import datetime
from io import BytesIO
from pathlib import Path
import tempfile
import uuid
import shutil

import main as m  # import the module (so we can override globals)

APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx"
VENDOR_INFO_PATH = APP_DIR / "รายงานข้อมูลผู้จำหน่าย.xlsx"  # if you keep it in repo

st.title("DONMARK Purchase Order Generator")
st.write("อัปโหลดไฟล์จาก Express และ รายละเอียดสินค้า แล้วเลือก supplier เพื่อสร้าง PO")

uploaded_update = st.file_uploader("Upload file from EXPRESS (.xlsx)", type=["xlsx"])
uploaded_a0029 = st.file_uploader("Upload รายละเอียดสินค้า （.xlsx)", type=["xlsx"])

vendor = st.text_input("Supplier code", value="A0029")
po_date = st.date_input("PO date", value=datetime.date.today())
rate = st.number_input("Exchange rate THB / CNY", value=6.0, step=0.1)

if st.button("Generate PO"):
    if uploaded_update is None:
        st.error("Please upload update_yuan.xlsx")
        st.stop()
    if uploaded_a0029 is None:
        st.error("Please upload A0029.xlsx")
        st.stop()
    if not vendor.strip():
        st.error("Please enter a vendor code.")
        st.stop()
    if not TEMPLATE_PATH.exists():
        st.error("Template file not found in app folder.")
        st.stop()

    run_id = uuid.uuid4().hex[:8]

    # Use a writable temp directory (Streamlit Cloud safe)
    work_dir = Path(tempfile.gettempdir()) / f"po_app_{run_id}"
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Save uploads into temp work folder
        update_path = work_dir / "update_yuan.xlsx"
        a0029_path = work_dir / "A0029.xlsx"

        update_path.write_bytes(uploaded_update.getbuffer())
        a0029_path.write_bytes(uploaded_a0029.getbuffer())

        # Per-run output folder
        output_buyers_dir = work_dir / "output_buyers"
        output_buyers_dir.mkdir(exist_ok=True)

        # Override globals in main.py so split_main() writes/reads from temp
        m.INPUT_XLSX = str(update_path)
        m.OUTPUT_FOLDER = str(output_buyers_dir)
        m.A0029_CATALOG_XLSX = str(a0029_path)

        # vendor info (optional)
        if VENDOR_INFO_PATH.exists():
            m.VENDOR_INFO_XLSX = str(VENDOR_INFO_PATH)

        st.info("Parsing update_yuan.xlsx and building buyer files...")
        m.main()  # your split function

        st.info("Generating PO...")
        buffer = BytesIO()
        m.generate_po_from_template(
            template_path=str(TEMPLATE_PATH),
            vendor_code=vendor.strip(),
            po_date=po_date,
            rate_thb_per_cny=rate,
            po_output_folder=None,
            output_stream=buffer,
            a0029_catalog_path=str(a0029_path),
        )

        buffer.seek(0)
        st.success("PO generated successfully!")
        st.download_button(
            label=f"Download PO_{vendor.strip()}.xlsx",
            data=buffer,
            file_name=f"PO_{vendor.strip()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    finally:
        # Clean up temp folder (optional)
        shutil.rmtree(work_dir, ignore_errors=True)
