import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2
import io

st.set_page_config(page_title="AI Chấm SKKN", layout="wide")

# Hàm đọc nội dung file
def doc_noi_dung(file):
    text = ""
    if file.name.endswith(".docx"):
        doc = docx.Document(file)
        for para in doc.paragraphs:
            text += para.text + "\n"
    elif file.name.endswith(".pdf"):
        pdf_reader = PyPDF2.PdfReader(file)
        for page in pdf_reader.pages:
            text += page.extract_text() + "\n"
    return text

# Hàm tạo file Word
def tao_file_word(ten_skkn, linh_vuc, ten_giam_khao, chuc_vu, noi_dung_ai):
    doc = docx.Document()

    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n").bold = True

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(
        "PHIẾU NHẬN XÉT, ĐÁNH GIÁ\nGIẢI PHÁP ĐỀ NGHỊ CÔNG NHẬN SÁNG KIẾN\n(Phiếu dành cho thành viên chuyên ngành)"
    )
    run_title.bold = True
    run_title.font.size = Pt(14)

    doc.add_paragraph(f"Họ và tên thành viên nhận xét: {ten_giam_khao}")
    doc.add_paragraph(f"Chức vụ: {chuc_vu}\n")

    doc.add_paragraph("NHẬN XÉT, ĐÁNH GIÁ GIẢI PHÁP").runs[0].bold = True
    doc.add_paragraph(f"1. Tên giải pháp: {ten_skkn}")
    doc.add_paragraph(f"2. Thuộc lĩnh vực: {linh_vuc}\n")

    doc.add_paragraph("3. Nhận xét và Chấm điểm (Do AI thực hiện dựa trên Mẫu 09):").runs[0].bold = True
    doc.add_paragraph(noi_dung_ai)

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sign.add_run("\nTHÀNH VIÊN NHẬN XÉT\n(Họ, tên và chữ ký)").bold = True

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- KHỞI TẠO SESSION STATE ---
if "linh_vuc_list" not in st.session_state:
    st.session_state.linh_vuc_list = ["Toán học", "Vật lí", "Hóa học", "Sinh học", "Ngữ văn", "Khác"]

# --- SIDEBAR ---
with st.sidebar:
    st.title("⚙️ Cấu hình Hệ thống")
    api_key = st.text_input("Nhập Gemini API Key:", type="password")
    model_chon = st.selectbox(
        "Chọn model AI:",
        [
            "gemini-3-flash-preview",
            "gemini-3.1-flash-lite",
            "gemini-3.1-pro-preview",
            "gemini-1.5-pro-latest",
            "gemini-pro-latest",
            "gemini-flash-latest",
            "gemini-flash-lite-latest",
        ],
    )
    st.info("Hệ thống tuân thủ thang điểm 100 theo Mẫu 09 và xuất form Mẫu 10.")

# --- PHẦN NHẬP THÔNG TIN ---
st.header("📝 Thông tin Định danh")
col1, col2 = st.columns(2)
with col1:
    ten_giam_khao = st.text_input("Tên thành viên nhận xét (Giám khảo):", value="Giám khảo AI")
    chuc_vu = st.text_input("Chức vụ của giám khảo:", value="Chuyên gia AI")
with col2:
    ten_skkn = st.text_input("Tên giải pháp/SKKN:")

    # Lĩnh vực với khả năng thêm mới
    st.write("**Lĩnh vực:**")
    col_lv, col_btn = st.columns([3, 1])
    with col_lv:
        linh_vuc = st.selectbox("Chọn lĩnh vực:", st.session_state.linh_vuc_list, label_visibility="collapsed")
    with col_btn:
        if st.button("＋ Thêm"):
            st.session_state.show_them_lv = True

    if st.session_state.get("show_them_lv"):
        new_lv = st.text_input("Nhập tên lĩnh vực mới:", key="new_lv_input")
        if st.button("Lưu lĩnh vực") and new_lv:
            if new_lv not in st.session_state.linh_vuc_list:
                st.session_state.linh_vuc_list.append(new_lv)
            st.session_state.show_them_lv = False
            st.rerun()

# --- PHẦN TẢI FILE ---
st.header("📄 Dữ liệu đầu vào")
uploaded_file = st.file_uploader("Tải lên file SKKN của bạn (Word/PDF)", type=["docx", "pdf"])

if st.button("Bắt đầu chấm điểm", type="primary"):
    if not api_key or not uploaded_file:
        st.error("Vui lòng nhập API Key và tải file lên!")
    else:
        with st.spinner("AI đang đọc tài liệu, đối chiếu thang điểm Mẫu 09 và viết nhận xét..."):
            try:
                noi_dung = doc_noi_dung(uploaded_file)
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(model_chon)

                prompt = f"""
                Bạn là thành viên hội đồng chuyên ngành chấm Sáng kiến kinh nghiệm. Hãy đọc nội dung SKKN sau và viết Phiếu nhận xét, đánh giá.

                BẮT BUỘC PHẢI CHẤM ĐIỂM THEO THANG 100 ĐIỂM SAU (Ghi rõ số điểm đạt được cho từng phần):
                1. Tính mới, tính sáng tạo: Tối đa 30 điểm. (Mức độ: Hoàn toàn mới 27-30đ, Cải tiến khá 21-26đ, Cải tiến trung bình 16-20đ, Cải tiến ít 1-15đ, Không có 0đ).
                2. Khả năng áp dụng: Tối đa 30 điểm. (Mức độ: Toàn tỉnh/ngoài tỉnh 27-30đ, Trong ngành/địa bàn tỉnh 21-26đ, Trong đơn vị 16-20đ, Ít khả năng 1-15đ, Không có 0đ).
                3. Hiệu quả của giải pháp: Tối đa 40 điểm. (Mức độ: Cao 31-40đ, Khá 21-30đ, Trung bình 11-20đ, Ít 1-10đ, Không 0đ).

                CẤU TRÚC CÂU TRẢ LỜI CỦA BẠN (Không dùng Markdown # hay * dư thừa):

                I. ĐÁNH GIÁ CHI TIẾT
                - Tính mới, tính sáng tạo (Điểm: .../30): Phân tích chi tiết tại sao cho mức điểm này.
                - Khả năng áp dụng (Điểm: .../30): Phân tích chi tiết.
                - Hiệu quả của giải pháp (Điểm: .../40): Phân tích chi tiết.

                II. TỔNG HỢP VÀ KẾT LUẬN
                - Tổng điểm: .../100
                - Điều kiện xét duyệt: Tổng điểm >= 70, Tính mới >= 20, Hiệu quả >= 25.
                - Nhận xét chung: (Đánh giá tóm tắt)
                - Kết luận: [Đạt / Không đạt] (Chỉ ghi Đạt nếu thỏa mãn CẢ 3 điều kiện trên).

                Nội dung SKKN:
                {noi_dung}
                """

                response = model.generate_content(prompt)
                st.session_state.ai_text = response.text
                st.session_state.word_bytes = tao_file_word(
                    ten_skkn, linh_vuc, ten_giam_khao, chuc_vu, response.text
                )
                st.success("Đã chấm xong!")

            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")

# --- HIỂN THỊ KẾT QUẢ ---
if "ai_text" in st.session_state:
    st.markdown("### 📊 Phiếu nhận xét (Bản xem trước)")
    st.write(st.session_state.ai_text)

    st.download_button(
        label="📥 TẢI XUỐNG FILE WORD (MẪU 10 + ĐIỂM SỐ MẪU 09)",
        data=st.session_state.word_bytes,
        file_name="Phieu_Danh_Gia_SKKN_Mau10.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
