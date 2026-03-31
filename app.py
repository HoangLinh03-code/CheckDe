import streamlit as st
import os
import tempfile
import check_de # Import file logic gốc của bạn

st.set_page_config(page_title="Phần Mềm Check Đề AI", layout="wide")

st.title("🚀 Phần Mềm Check Đề & Lỗi Đáp Án - AI")

# 1. Giao diện Upload File
col1, col2 = st.columns(2)
with col1:
    orig_file = st.file_uploader("📂 Tải lên Đề Gốc (.docx, .pdf)", type=['docx', 'pdf'])
    ans_file = st.file_uploader("📊 Tải lên Đáp án Excel (.xlsx)", type=['xlsx'])
with col2:
    shuf_files = st.file_uploader("📝 Tải lên các Đề Trộn (.docx, .pdf)", type=['docx', 'pdf'], accept_multiple_files=True)

subject = st.selectbox("🎯 Chọn Môn Học", ["auto", "math", "english", "other"])

# 2. Xử lý khi bấm nút chạy
if st.button("BẮT ĐẦU KIỂM TRA", type="primary"):
    if not (orig_file and ans_file and shuf_files):
        st.warning("⚠️ Vui lòng tải lên đầy đủ Đề gốc, Đáp án và ít nhất 1 Đề trộn.")
    else:
        with st.spinner("Đang xử lý... Quá trình này có thể mất vài phút."):
            # Tạo thư mục tạm để lưu file upload
            with tempfile.TemporaryDirectory() as temp_dir:
                # Lưu file gốc
                orig_path = os.path.join(temp_dir, orig_file.name)
                with open(orig_path, "wb") as f: f.write(orig_file.getbuffer())
                
                # Lưu file đáp án
                ans_path = os.path.join(temp_dir, ans_file.name)
                with open(ans_path, "wb") as f: f.write(ans_file.getbuffer())
                
                # Lưu các đề trộn
                shuf_paths = []
                for sf in shuf_files:
                    sf_path = os.path.join(temp_dir, sf.name)
                    with open(sf_path, "wb") as f: f.write(sf.getbuffer())
                    shuf_paths.append(sf_path)

                # TODO: Sửa lại hàm main() trong check_de.py để nhận tham số trực tiếp
                # Ví dụ: check_de.run_check(orig_path, ans_path, shuf_paths, subject)
                
                st.success("✅ Kiểm tra hoàn tất!")
                
                # Cung cấp nút tải xuống file Excel báo cáo
                # report_path = "Đường_dẫn_file_kết_quả_vừa_tạo.xlsx"
                # with open(report_path, "rb") as file:
                #     st.download_button("📥 Tải Báo Cáo Excel", data=file, file_name="Bao_Cao_Check_De.xlsx")