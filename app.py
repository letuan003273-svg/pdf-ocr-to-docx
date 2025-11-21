import streamlit as st
import tempfile
import io
import zipfile
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from PyPDF2 import PdfReader

st.set_page_config(page_title="OCR & PDF Extract → TEXT", layout="centered")

st.title("Chuyển đổi PDF / Ảnh sang Văn bản (TEXT)")
st.write("Hỗ trợ PDF, PNG, JPG, JPEG. Tự động OCR nếu là file ảnh hoặc PDF scan.")

uploaded_files = st.file_uploader(
    "Chọn file để xử lý:",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

zip_option = st.checkbox("Nén tất cả file kết quả vào ZIP", value=True)
process_btn = st.button("Bắt đầu chuyển đổi")


# ===============================
# 🔧 HÀM XỬ LÝ TỪNG FILE
# ===============================

def extract_text_from_pdf(pdf_path: Path) -> str:
    """
    Cố gắng lấy text trực tiếp từ PDF.
    Nếu không có text (scanned PDF) → fallback OCR.
    """
    text = ""

    # Thử trích text trực tiếp
    try:
        reader = PdfReader(str(pdf_path))
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    except:
        pass

    if text.strip():
        return text

    # Fallback OCR nếu PDF không có text (scanned PDF)
    try:
        pages = convert_from_path(str(pdf_path))
        text_ocr = ""
        for pg in pages:
            text_ocr += pytesseract.image_to_string(pg) + "\n"
        return text_ocr
    except Exception as e:
        return f"[LỖI OCR PDF]: {e}"


def extract_text_from_image(img_path: Path) -> str:
    try:
        img = Image.open(str(img_path))
        return pytesseract.image_to_string(img)
    except Exception as e:
        return f"[LỖI OCR ẢNH]: {e}"


def process_file(input_path: Path) -> str:
    """
    Trả về text của file.
    """
    ext = input_path.suffix.lower()

    if ext == ".pdf":
        return extract_text_from_pdf(input_path)

    elif ext in [".png", ".jpg", ".jpeg"]:
        return extract_text_from_image(input_path)

    else:
        return "[Định dạng không hỗ trợ]"


# ===============================
# ▶️ BẮT ĐẦU XỬ LÝ
# ===============================

if process_btn:
    if not uploaded_files:
        st.warning("Vui lòng tải ít nhất 1 file.")
    else:
        st.subheader("Kết quả:")

        results = []   # (filename, text)

        progress = st.progress(0)
        total = len(uploaded_files)

        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            for idx, uf in enumerate(uploaded_files, start=1):

                save_path = tmpdir / uf.name
                with open(save_path, "wb") as f:
                    f.write(uf.read())

                text_result = process_file(save_path)
                results.append((uf.name, text_result))

                progress.progress(int(idx / total * 100))

        # Hiển thị + nút tải từng file
        for filename, text_content in results:
            st.markdown(f"### 📄 {filename}")
            st.text_area("Nội dung trích xuất:", text_content, height=200)

            st.download_button(
                label=f"Tải xuống {filename}.txt",
                data=text_content,
                file_name=f"{Path(filename).stem}.txt",
                mime="text/plain"
            )

        # ZIP tất cả
        if zip_option:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for filename, text in results:
                    zf.writestr(f"{Path(filename).stem}.txt", text)

            st.download_button(
                label="📦 Tải về tất cả file TEXT (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="converted_texts.zip",
                mime="application/zip"
            )
