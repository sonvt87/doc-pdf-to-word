import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from docx import Document
from docx.shared import Pt
import io

# Cấu hình Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Giao diện
st.set_page_config(page_title="PDF to Word", page_icon="📄")
st.title("📄 Chuyển PDF sang Word - Sản phẩm thử nghiệm của Sơn Vũ")
st.write("Tải file PDF, tôi sẽ giúp bạn trích toàn bộ văn bản và lưu thành file Word Times New Roman, cỡ 14 ✨")

# Upload file
uploaded_file = st.file_uploader("📤 Tải file PDF lên", type="pdf")

# OCR và trích văn bản
def extract_text_with_ocr(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    full_text = ""
    for page_num, page in enumerate(doc):
        text = page.get_text("text")
        if text.strip():
            full_text += text.strip() + "\n\n"
        else:
            # Khi không có văn bản, sử dụng OCR
            pix = page.get_pixmap(dpi=300)  # Tăng độ phân giải để cải thiện OCR
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            text_from_image = pytesseract.image_to_string(img, lang='vie')
            full_text += f"[Trang {page_num + 1} - OCR]\n{text_from_image.strip()}\n\n"
    return full_text

# Tạo file Word
def create_docx(text):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)  # Cỡ chữ 14 theo yêu cầu

    for para in text.strip().split("\n\n"):
        doc.add_paragraph(para.strip())

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Xử lý khi có file
if uploaded_file:
    with st.spinner("🔍 Đang trích xuất văn bản từ PDF..."):
        try:
            raw_text = extract_text_with_ocr(uploaded_file)
            st.success("✅ Trích xuất văn bản thành công!")
        except Exception as e:
            st.error(f"❌ Lỗi khi trích xuất văn bản: {e}")
            raw_text = ""

        if raw_text:
            # Hiển thị văn bản gốc trong một cửa sổ
            st.text_area("📃 Văn bản gốc", raw_text, height=300, key="raw_text_display")

            # Tạo file Word và cho phép tải xuống
            docx_file = create_docx(raw_text)
            st.download_button(
                label="📥 Tải xuống file Word (.docx) chuẩn",
                data=docx_file,
                file_name="Van_ban_chuan.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
