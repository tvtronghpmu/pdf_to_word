import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from docx import Document
import os
import io

# --- CÀI ĐẶT BAN ĐẦU ---

# --- Cấu hình đường dẫn Tesseract (QUAN TRỌNG) ---
# Dòng này rất quan trọng nếu bạn cài Tesseract ở một vị trí không chuẩn.
# Ví dụ trên Windows:
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# Trên macOS/Linux, thường thì không cần nếu bạn đã thêm vào PATH.
# Hãy chắc chắn rằng bạn đã cài đặt Tesseract-OCR trên máy của mình.
# Hướng dẫn cài đặt: https://github.com/tesseract-ocr/tesseract

# --- Tạo thư mục lưu file output ---
output_folder = "converted_docs"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# --- HÀM XỬ LÝ CHÍNH ---

def convert_pdf_to_word(uploaded_file):
    """
    Hàm chính để chuyển đổi file PDF được tải lên sang file Word.
    Hàm sẽ xử lý cả văn bản thuần túy và hình ảnh (sử dụng OCR).
    """
    try:
        # Mở file PDF từ dữ liệu trong bộ nhớ
        pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # Tạo một tài liệu Word mới
        doc = Document()

        # Lặp qua từng trang của file PDF
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            
            # --- 1. Trích xuất văn bản thuần túy ---
            text = page.get_text("text")
            if text.strip():  # Nếu có văn bản, thêm vào file Word
                doc.add_paragraph(text)

            # --- 2. Trích xuất hình ảnh và sử dụng OCR ---
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                
                # Chuyển đổi bytes hình ảnh sang đối tượng Image của Pillow
                image = Image.open(io.BytesIO(image_bytes))

                # Sử dụng Tesseract để nhận diện văn bản từ hình ảnh
                # Chỉ định ngôn ngữ là tiếng Việt + tiếng Anh để tăng độ chính xác
                ocr_text = pytesseract.image_to_string(image, lang='vie+eng')
                
                if ocr_text.strip():
                    doc.add_paragraph("[Văn bản từ hình ảnh (OCR)]:")
                    doc.add_paragraph(ocr_text)

            # Thêm dấu ngắt trang để giữ nguyên bố cục tương đối
            if page_num < len(pdf_document) - 1:
                doc.add_page_break()

        # Tạo đường dẫn lưu file output
        output_filename = os.path.splitext(uploaded_file.name)[0] + ".docx"
        output_path = os.path.join(output_folder, output_filename)
        
        # Lưu tài liệu Word
        doc.save(output_path)
        
        return output_path

    except Exception as e:
        st.error(f"Đã xảy ra lỗi: {e}")
        return None

# --- GIAO DIỆN NGƯỜI DÙNG VỚI STREAMLIT ---

# --- Tiêu đề của ứng dụng ---
st.title("Chuyển đổi PDF sang Word (với OCR) 📄➡️📝")
st.write("Tải lên file PDF của bạn, ứng dụng sẽ chuyển đổi nó sang định dạng .docx.")
st.write("Nếu PDF chứa hình ảnh, công nghệ OCR (Nhận dạng ký tự quang học) sẽ được sử dụng để trích xuất văn bản từ những hình ảnh đó.")

# --- Mục tải file lên ---
uploaded_file = st.file_uploader("Chọn một file PDF", type=["pdf"])

# --- Nút thực hiện chuyển đổi ---
if st.button("Chuyển đổi"):
    if uploaded_file is not None:
        with st.spinner('Đang xử lý... Quá trình này có thể mất vài phút tùy thuộc vào file PDF.'):
            # Gọi hàm chuyển đổi
            result_path = convert_pdf_to_word(uploaded_file)
        
        if result_path:
            st.success("Chuyển đổi thành công!")
            st.info(f"File đã được lưu tại thư mục cục bộ: `{result_path}`")
            
            # Cung cấp nút để tải file đã chuyển đổi về trình duyệt
            with open(result_path, "rb") as f:
                st.download_button(
                    label="Tải file Word (.docx)",
                    data=f,
                    file_name=os.path.basename(result_path),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.warning("Vui lòng tải lên một file PDF trước.")

# --- Chú thích thêm ---
st.markdown("---")
