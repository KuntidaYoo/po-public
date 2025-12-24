import streamlit as st
import datetime
from io import BytesIO
from pathlib import Path
import tempfile
import uuid
import shutil

import main as m  # import the module so we can override globals per-run


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx"

st.set_page_config(
    page_title="DONMARK PO Generator",
    page_icon="🧾",
    layout="centered",
)

st.title("DONMARK Purchase Order Generator")
st.write("อัปโหลดไฟล์จาก Express + รายละเอียดสินค้า + ข้อมูลผู้จำหน่าย แล้วเลือก Supplier เพื่อสร้าง PO")

# --- Uploads ---
uploaded_update = st.file_uploader("1) Upload file from EXPRESS (update_yuan.xlsx)", type=["xlsx"])
uploaded_a0029 = st.file_uploader("2) Upload รายละเอียดสินค้า (A0029.xlsx)", type=["xlsx"])
uploaded_vendor_info = st.file_uploader("3) Upload รายงานข้อมูลผู้จำหน่าย.xlsx (Vendor info)", type=["xlsx"])

st.divider()

# --- Inputs ---
vendor = st.text_input("Supplier code (เช่น A0029)", value="A0029")
po_date = st.date_input("PO date", value=datetime.date.today())
rate = st.number_input("Exchange rate THB / CNY", value=6.0, step=0.1)


def _save_upload(uploaded, path: Path):
    """Write Streamlit uploaded file to disk."""
    path.write_bytes(uploaded.getbuffer())


if st.button("Generate PO", type="primary"):
    # --- Validate ---
    if uploaded_update is None:
        st.error("Please upload update_yuan.xlsx")
        st.stop()
    if uploaded_a0029 is None:
        st.error("Please upload A0029.xlsx")
        st.stop()
    if uploaded_vendor_info is None:
        st.error("Please upload รายงานข้อมูลผู้จำหน่าย.xlsx")
        st.stop()

    vendor_code = vendor.strip().upper()
    if not vendor_code:
        st.error("Please enter a Supplier code (Vendor code).")
        st.stop()

    if not TEMPLATE_PATH.exists():
        st.error("Template file not found: ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx (must be in the same folder as app.py)")
        st.stop()

    run_id = uuid.uuid4().hex[:8]
    work_dir = Path(tempfile.gettempdir()) / f"po_app_{run_id}"
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        # --- Save uploads into temp folder ---
        update_path = work_dir / "update_yuan.xlsx"
        a0029_path = work_dir / "A0029.xlsx"
        vendor_info_path = work_dir / "รายงานข้อมูลผู้จำหน่าย.xlsx"

        _save_upload(uploaded_update, update_path)
        _save_upload(uploaded_a0029, a0029_path)
        _save_upload(uploaded_vendor_info, vendor_info_path)

        # --- Output folder (temp) ---
        output_buyers_dir = work_dir / "output_buyers"
        output_buyers_dir.mkdir(exist_ok=True)

        # --- Override globals in main.py so it reads/writes in temp ---
        m.INPUT_XLSX = str(update_path)
        m.OUTPUT_FOLDER = str(output_buyers_dir)
        m.A0029_CATALOG_XLSX = str(a0029_path)
        m.VENDOR_INFO_XLSX = str(vendor_info_path)  # ✅ ALWAYS use uploaded vendor info

        # --- Run pipeline ---
        st.info("Parsing EXPRESS file and building buyer files...")
        m.main()

        st.info("Generating PO...")
        buffer = BytesIO()
        m.generate_po_from_template(
            template_path=str(TEMPLATE_PATH),
            vendor_code=vendor_code,
            po_date=po_date,
            rate_thb_per_cny=rate,
            po_output_folder=None,
            output_stream=buffer,
            a0029_catalog_path=str(a0029_path),
        )
        buffer.seek(0)

        st.success("PO generated successfully!")
        st.download_button(
            label=f"Download PO_{vendor_code}.xlsx",
            data=buffer,
            file_name=f"PO_{vendor_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error("Error while generating PO:")
        st.exception(e)

    finally:
        # Clean up temp folder
        shutil.rmtree(work_dir, ignore_errors=True)
