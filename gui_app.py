# -*- coding: utf-8 -*-
import sys
import os
import re
import shutil
import pythoncom
from docx2pdf import convert
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QFileDialog, QListWidget,
                             QComboBox, QGroupBox, QTextEdit, QMessageBox, QDialog, QFormLayout, QCheckBox,
                             QSplitter, QPlainTextEdit, QToolBar, QAction) # Bổ sung thêm các module này
from PyQt5.QtGui import QFont, QIcon, QKeySequence
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject

# =========================================================
# VÁ LỖI PYINSTALLER: ĐƯA IMPORT RA GLOBAL ĐỂ GÓI KÈM CODE
# =========================================================
import check_de
import prompt_loader
import math_exam_handler
import callAPI

# ---------------------------------------------------------
# HÀM LẤY ĐƯỜNG DẪN GỐC (Hỗ trợ khi đóng gói PyInstaller)
# ---------------------------------------------------------
def get_app_dir():
    """Lấy thư mục gốc chứa script hoặc file .exe"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def setup_environment():
    """Khởi tạo môi trường, giải quyết xung đột đường dẫn khi đóng gói .exe"""
    user_prompts_dir = os.path.join(get_app_dir(), 'prompts')
    
    # Nếu chạy .exe và thư mục prompts chưa tồn tại ở ngoài, copy từ file build ra
    if getattr(sys, 'frozen', False) and not os.path.exists(user_prompts_dir):
        os.makedirs(user_prompts_dir, exist_ok=True)
        bundled_prompts_dir = os.path.join(sys._MEIPASS, 'prompts')
        if os.path.exists(bundled_prompts_dir):
            for file in os.listdir(bundled_prompts_dir):
                shutil.copy2(os.path.join(bundled_prompts_dir, file), os.path.join(user_prompts_dir, file))
    
    # Ép prompt_loader phải đọc/ghi vào thư mục prompts nằm cạnh file .exe
    prompt_loader._PROMPTS_DIR = user_prompts_dir

# Khởi chạy setup ngay lập tức
setup_environment()

# ---------------------------------------------------------
# LỚP ĐIỀU HƯỚNG CONSOLE LOG LÊN GIAO DIỆN
# ---------------------------------------------------------
class EmittingStream(QObject):
    textWritten = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.encoding = 'utf-8'

    def write(self, text):
        self.textWritten.emit(str(text))
        
    def flush(self):
        pass
        
    def reconfigure(self, **kwargs):
        if 'encoding' in kwargs:
            self.encoding = kwargs['encoding']

# ---------------------------------------------------------
# WORKER THREAD: Xử lý logic ngầm (Chuyển PDF + Chạy Check)
# ---------------------------------------------------------
class CheckWorker(QThread):
    finished_signal = pyqtSignal(str)
    error_signal = pyqtSignal(str)
    log_signal = pyqtSignal(str)

    def __init__(self, original_file, answer_file, shuffled_files, subject, vision_mode, auto_pdf):
        super().__init__()
        self.original_file = original_file
        self.answer_file = answer_file
        self.shuffled_files = shuffled_files
        self.subject = subject
        self.vision_mode = vision_mode
        self.auto_pdf = auto_pdf

    def convert_to_pdf_if_needed(self, filepath):
        if not filepath.lower().endswith('.docx'):
            return filepath
            
        pdf_path = filepath[:-5] + '.pdf'
        if not os.path.exists(pdf_path):
            self.log_signal.emit(f"  🔄 Đang tự động chuyển đổi sang PDF để giữ nguyên công thức: {os.path.basename(filepath)}...")
            try:
                convert(filepath, pdf_path)
                self.log_signal.emit(f"  ✅ Chuyển đổi thành công: {os.path.basename(pdf_path)}")
            except Exception as e:
                self.log_signal.emit(f"  ❌ Lỗi khi convert PDF (Bỏ qua convert): {e}")
                return filepath
        else:
            self.log_signal.emit(f"  ⚡ Đã có sẵn bản PDF, sử dụng file: {os.path.basename(pdf_path)}")
        return pdf_path

    def run(self):
        try:
            # QUAN TRỌNG: Bao bọc COM object bằng try...finally để chống treo App
            pythoncom.CoInitialize()
            try:
                actual_orig_file = self.original_file
                actual_shuffled_files = list(self.shuffled_files)
                
                if self.auto_pdf or self.vision_mode:
                    if self.auto_pdf:
                        self.log_signal.emit("\n⚙️ Chế độ Môn Khoa Học/Toán: Kích hoạt tự động chuyển DOCX -> PDF.")
                    
                    self.convert_to_pdf_if_needed(self.original_file)
                    for f in actual_shuffled_files:
                        self.convert_to_pdf_if_needed(f)
                    
                    self.vision_mode = True
            finally:
                pythoncom.CoUninitialize()

            # Bắt đầu chạy Check Đề
            sys.argv = ['check_de.py', 
                        '-o', actual_orig_file, 
                        '-a', self.answer_file, 
                        '-f'] + actual_shuffled_files
            
            if self.subject != 'auto':
                sys.argv.extend(['--subject', self.subject])
            if self.vision_mode:
                sys.argv.append('--vision')

            self.log_signal.emit("\n🚀 Bắt đầu phân tích AI...")
            check_de.main()
            self.finished_signal.emit("✅ Hoàn tất kiểm tra đề!")
            
        except Exception as e:
            self.error_signal.emit(f"❌ Lỗi: {str(e)}")

# ---------------------------------------------------------
# DIALOG: THÊM VÀ TÙY CHỈNH PROMPT
# ---------------------------------------------------------
class AddSubjectDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Thêm / Tùy Chỉnh Prompt Môn Học")
        self.resize(700, 600)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.txt_id = QLineEdit()
        self.txt_id.setPlaceholderText("VD: vatly (viết liền không dấu)")
        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("VD: Vật Lý")
        
        self.cb_type = QComboBox()
        self.cb_type.addItem("Khoa học Tự nhiên (KHTN - Có công thức, tự động chuyển PDF)", "khtn")
        self.cb_type.addItem("Khoa học Xã hội (KHXH - Chỉ có text, đọc trực tiếp DOCX)", "khxh")
        
        form.addRow("Mã môn (ID):", self.txt_id)
        form.addRow("Tên hiển thị:", self.txt_name)
        form.addRow("Đặc thù môn:", self.cb_type)
        layout.addLayout(form)

        btn_generate = QPushButton("🔄 Sinh Template Mẫu (Dựa trên Đặc thù môn)")
        btn_generate.setStyleSheet("background-color: #ff9800; color: white; font-weight: bold;")
        btn_generate.clicked.connect(self.generate_template)
        layout.addWidget(btn_generate)

        layout.addWidget(QLabel("<b>Nội dung Prompt (Bạn có thể tùy chỉnh tự do trước khi Lưu):</b>"))
        self.txt_prompt = QTextEdit()
        self.txt_prompt.setPlaceholderText("Bấm 'Sinh Template Mẫu' để tạo khung hoặc tự dán nội dung prompt của bạn vào đây...")
        layout.addWidget(self.txt_prompt)

        btn_layout = QHBoxLayout()
        btn_save = QPushButton("Lưu Prompt")
        btn_save.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 5px 15px;")
        btn_save.clicked.connect(self.save_prompt)
        btn_cancel = QPushButton("Hủy")
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def generate_template(self):
        subj_id = self.txt_id.text().strip().lower() or "mon_hoc"
        subj_name = self.txt_name.text().strip() or "Môn Học"
        subj_type = self.cb_type.currentData()
        
        if subj_type == "khtn":
            template = f"""[SUBJECT]\n{subj_id}\n\n[SUBJECT_LABEL]\n{subj_name}\n\n[AUTO_PDF]\nTrue\n\n[MAX_PAGES]\n4\n\n[PARTS_CONFIG]\n1|mc|Trắc nghiệm nhiều phương án|nhiều phương án,trắc nghiệm nhiều\n2|tf|Trắc nghiệm Đúng/Sai|đúng sai,đúng hoặc sai\n3|fill|Trả lời ngắn / Điền số|trả lời ngắn,điền,ngắn\n\n[PART_LABELS]\n1=PHẦN I — Trắc nghiệm nhiều phương án\n2=PHẦN II — Trắc nghiệm Đúng/Sai\n3=PHẦN III — Trả lời ngắn / Điền số\n\n[PROMPT_HEADER]\nBạn là chuyên gia kiểm tra đề thi {subj_name}. Nhiệm vụ: Khớp các câu hỏi của ĐỀ TRỘN (mã {{exam_code}}) vào đúng vị trí ĐỀ GỐC.\nCẤU TRÚC PHÂN VÙNG BẮT BUỘC TUÂN THỦ:\n1. Số lượng câu hỏi mỗi phần là LINH HOẠT, phụ thuộc hoàn toàn vào Đề gốc.\n2. "original_q" PHẢI là số thứ tự câu hỏi ĐÃ ĐƯỢC CỘNG DỒN (Global Number) tính từ câu đầu tiên của Đề gốc.\n3. PHẦN IV (Tự luận): BỎ QUA HOÀN TOÀN.\n\n[OUTPUT_FORMAT]\nTrả về JSON array. Bắt buộc có "part", "q_type" và "original_q" (theo Global Number).\nCHỈ TRẢ VỀ JSON ARRAY.\nKHÔNG GIẢI THÍCH.\n\n[VISION_PROMPT]\nBạn là chuyên gia kiểm tra đề thi {subj_name}. Tôi đính kèm 2 file: ĐỀ GỐC và ĐỀ TRỘN (mã {{exam_code}}).\nNHIỆM VỤ: Đọc cả 2 file, khớp các câu hỏi vào vị trí tương ứng. Số lượng câu mỗi phần là LINH HOẠT. "original_q" phải là số thứ tự cộng dồn từ câu 1 của đề gốc. Đề có 3 PHẦN: MC, TF, Fill.\nTrả về JSON array."""
        else:
            template = f"""[SUBJECT]\n{subj_id}\n\n[SUBJECT_LABEL]\n{subj_name}\n\n[AUTO_PDF]\nFalse\n\n[MAX_PAGES]\n6\n\n[PARTS_CONFIG]\n1|mc|Trắc nghiệm|chọn đáp án đúng\n\n[PART_LABELS]\n1=Trắc nghiệm (A/B/C/D)\n\n[PROMPT_HEADER]\nBạn là chuyên gia kiểm tra đề thi {subj_name}. Nhiệm vụ: khớp từng câu hỏi trong ĐỀ TRỘN (mã {{exam_code}}) với câu hỏi tương ứng trong ĐỀ GỐC.\nLƯU Ý: Đề {subj_name} CHỈ CÓ TRẮC NGHIỆM.\n\n[OUTPUT_FORMAT]\nTrả về JSON array, mỗi phần tử có:\n{{"shuffled_q": N, "original_q": N, "part": 1, "q_type": "mc", "option_mapping": {{...}}}}\nCHỈ TRẢ VỀ JSON ARRAY.\nKHÔNG GIẢI THÍCH.\n\n[VISION_PROMPT]\nBạn là chuyên gia kiểm tra đề thi {subj_name}. Tôi đính kèm 2 file: ĐỀ GỐC và ĐỀ TRỘN (mã {{exam_code}}).\nNHIỆM VỤ: Đọc cả 2 file, khớp từng câu hỏi đề trộn với câu tương ứng trong đề gốc. CHỈ TRẢ VỀ JSON ARRAY."""
            
        self.txt_prompt.setPlainText(template)

    def save_prompt(self):
        subj_id = self.txt_id.text().strip().lower()
        if not subj_id or not subj_id.isalnum():
            QMessageBox.warning(self, "Lỗi", "Mã môn chỉ được chứa chữ cái và số, không khoảng trắng.")
            return
            
        content = self.txt_prompt.toPlainText().strip()
        if not content:
            QMessageBox.warning(self, "Lỗi", "Nội dung Prompt không được để trống.")
            return

        prompt_dir = os.path.join(get_app_dir(), 'prompts')
        os.makedirs(prompt_dir, exist_ok=True)
        
        file_path = os.path.join(prompt_dir, f"prompt_{subj_id}.txt")
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            QMessageBox.information(self, "Thành công", f"Đã lưu thành công Prompt cho môn: {self.txt_name.text()}")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file: {str(e)}")

# ---------------------------------------------------------
# DIALOG: QUẢN LÝ VÀ CHỈNH SỬA PROMPT HIỆN CÓ
# ---------------------------------------------------------
class PromptManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Quản Lý & Chỉnh Sửa Prompt")
        self.resize(1000, 700)
        self.current_file = None
        self.setup_ui()
        self.load_prompt_list()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Dùng QSplitter để chia đôi màn hình: Trái (Danh sách) - Phải (Editor)
        splitter = QSplitter(Qt.Horizontal)
        
        # --- BÊN TRÁI: DANH SÁCH PROMPT ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_layout.addWidget(QLabel("<b>Danh sách Prompt:</b>"))
        self.list_prompts = QListWidget()
        self.list_prompts.itemClicked.connect(self.on_file_selected)
        left_layout.addWidget(self.list_prompts)
        
        btn_new = QPushButton("✨ Tạo Mới (Template)")
        btn_new.clicked.connect(self.create_new_prompt)
        left_layout.addWidget(btn_new)
        
        # --- BÊN PHẢI: KHU VỰC EDITOR ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.lbl_current_file = QLabel("<i>Chưa chọn file nào</i>")
        right_layout.addWidget(self.lbl_current_file)
        
        # Thanh công cụ (Toolbar)
        toolbar = QToolBar()
        right_layout.addWidget(toolbar)
        
        # Các nút trên Toolbar
        btn_save = QAction("💾 Lưu Lại (Ctrl+S)", self)
        btn_save.setShortcut(QKeySequence("Ctrl+S"))
        btn_save.triggered.connect(self.save_current_prompt)
        toolbar.addAction(btn_save)
        
        toolbar.addSeparator()
        
        btn_undo = QAction("↩️ Undo", self)
        btn_undo.setShortcut(QKeySequence("Ctrl+Z"))
        btn_undo.triggered.connect(lambda: self.editor.undo())
        toolbar.addAction(btn_undo)
        
        btn_redo = QAction("↪️ Redo", self)
        btn_redo.setShortcut(QKeySequence("Ctrl+Y"))
        btn_redo.triggered.connect(lambda: self.editor.redo())
        toolbar.addAction(btn_redo)
        
        toolbar.addSeparator()
        
        btn_zoom_in = QAction("🔍 Tăng Font", self)
        btn_zoom_in.triggered.connect(self.zoom_in)
        toolbar.addAction(btn_zoom_in)
        
        btn_zoom_out = QAction("🔍 Giảm Font", self)
        btn_zoom_out.triggered.connect(self.zoom_out)
        toolbar.addAction(btn_zoom_out)

        toolbar.addSeparator()

        # Cụm chèn nhanh Tag
        btn_tag_header = QAction("🏷️ Chèn [PROMPT_HEADER]", self)
        btn_tag_header.triggered.connect(lambda: self.insert_text("\n[PROMPT_HEADER]\n"))
        toolbar.addAction(btn_tag_header)

        btn_tag_vision = QAction("🏷️ Chèn [VISION_PROMPT]", self)
        btn_tag_vision.triggered.connect(lambda: self.insert_text("\n[VISION_PROMPT]\n"))
        toolbar.addAction(btn_tag_vision)
        
        # Khung soạn thảo (Dùng QPlainTextEdit để tối ưu text thuần)
        self.editor = QPlainTextEdit()
        font = QFont("Consolas", 11)
        self.editor.setFont(font)
        self.editor.setPlaceholderText("Chọn một prompt bên trái để bắt đầu chỉnh sửa...")
        right_layout.addWidget(self.editor)
        
        # Thêm vào Splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([250, 750]) # Tỷ lệ chia màn hình ban đầu
        
        layout.addWidget(splitter)

    def load_prompt_list(self):
        self.list_prompts.clear()
        prompt_dir = os.path.join(get_app_dir(), 'prompts')
        if os.path.exists(prompt_dir):
            for file in os.listdir(prompt_dir):
                if file.startswith("prompt_") and file.endswith(".txt"):
                    self.list_prompts.addItem(file)

    def on_file_selected(self, item):
        filename = item.text()
        filepath = os.path.join(get_app_dir(), 'prompts', filename)
        self.current_file = filepath
        self.lbl_current_file.setText(f"<b>Đang sửa:</b> {filename}")
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            self.editor.setPlainText(content)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể đọc file: {e}")

    def save_current_prompt(self):
        if not self.current_file:
            QMessageBox.warning(self, "Cảnh báo", "Bạn chưa chọn file nào để lưu!")
            return
            
        content = self.editor.toPlainText()
        try:
            with open(self.current_file, 'w', encoding='utf-8') as f:
                f.write(content)
            QMessageBox.information(self, "Thành công", f"Đã lưu thành công:\n{os.path.basename(self.current_file)}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file: {e}")

    def create_new_prompt(self):
        # Mở dialog tạo mới cũ của bạn
        dialog = AddSubjectDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_prompt_list()

    def insert_text(self, text):
        self.editor.insertPlainText(text)

    def zoom_in(self):
        font = self.editor.font()
        font.setPointSize(font.pointSize() + 1)
        self.editor.setFont(font)

    def zoom_out(self):
        font = self.editor.font()
        if font.pointSize() > 6:
            font.setPointSize(font.pointSize() - 1)
            self.editor.setFont(font)


# ---------------------------------------------------------
# GIAO DIỆN CHÍNH
# ---------------------------------------------------------
class MainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Phần Mềm Check Đề & Lỗi Đáp Án - AI")
        self.resize(900, 750)
        self.shuffled_files_list = []
        self.setup_ui()
        self.load_subjects()
        
        sys.stdout = EmittingStream()
        sys.stdout.textWritten.connect(self.normal_output_written)
        sys.stderr = EmittingStream()
        sys.stderr.textWritten.connect(self.normal_output_written)

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        grp_input = QGroupBox("1. Chọn File Đầu Vào")
        vbox_input = QVBoxLayout(grp_input)

        hbox_orig = QHBoxLayout()
        self.lbl_orig = QLineEdit()
        self.lbl_orig.setReadOnly(True)
        self.lbl_orig.setPlaceholderText("Chưa chọn đề gốc (.docx, .pdf)...")
        btn_orig = QPushButton("Chọn Đề Gốc")
        btn_orig.clicked.connect(self.browse_original)
        hbox_orig.addWidget(self.lbl_orig)
        hbox_orig.addWidget(btn_orig)
        vbox_input.addLayout(hbox_orig)

        hbox_ans = QHBoxLayout()
        self.lbl_ans = QLineEdit()
        self.lbl_ans.setReadOnly(True)
        self.lbl_ans.setPlaceholderText("Chưa chọn file đáp án (.xlsx)...")
        btn_ans = QPushButton("Chọn Đáp Án Excel")
        btn_ans.clicked.connect(self.browse_answer)
        hbox_ans.addWidget(self.lbl_ans)
        hbox_ans.addWidget(btn_ans)
        vbox_input.addLayout(hbox_ans)

        hbox_shuf_btn = QHBoxLayout()
        hbox_shuf_btn.addWidget(QLabel("Danh sách đề trộn:"))
        btn_add_shuf = QPushButton("+ Thêm Đề Trộn")
        btn_add_shuf.clicked.connect(self.browse_shuffled)
        btn_clear_shuf = QPushButton("- Xóa Danh Sách")
        btn_clear_shuf.clicked.connect(self.clear_shuffled)
        hbox_shuf_btn.addStretch()
        hbox_shuf_btn.addWidget(btn_add_shuf)
        hbox_shuf_btn.addWidget(btn_clear_shuf)
        vbox_input.addLayout(hbox_shuf_btn)

        self.list_shuffled = QListWidget()
        vbox_input.addWidget(self.list_shuffled)
        main_layout.addWidget(grp_input)

        grp_settings = QGroupBox("2. Cấu Hình Kiểm Tra")
        hbox_settings = QHBoxLayout(grp_settings)
        
        hbox_settings.addWidget(QLabel("Môn học:"))
        self.cb_subject = QComboBox()
        hbox_settings.addWidget(self.cb_subject)
        
        btn_manage_prompt = QPushButton("📝 Quản Lý & Sửa Prompt")
        btn_manage_prompt.setStyleSheet("background-color: #607D8B; color: white; font-weight: bold;")
        btn_manage_prompt.clicked.connect(self.open_prompt_manager)
        hbox_settings.addWidget(btn_manage_prompt)

        self.chk_vision = QCheckBox("Sử dụng Vision Mode")
        hbox_settings.addWidget(self.chk_vision)
        
        hbox_settings.addStretch()
        main_layout.addWidget(grp_settings)

        self.btn_run = QPushButton("🚀 BẮT ĐẦU KIỂM TRA ĐỀ")
        self.btn_run.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        self.btn_run.clicked.connect(self.start_checking)
        main_layout.addWidget(self.btn_run)

        grp_log = QGroupBox("Log Trạng Thái")
        vbox_log = QVBoxLayout(grp_log)
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background-color: #1e1e1e; color: #00ff00; font-family: Consolas;")
        vbox_log.addWidget(self.txt_log)
        main_layout.addWidget(grp_log)

    def load_subjects(self):
        self.cb_subject.clear()
        self.cb_subject.addItem("Tự động nhận diện (Auto)", "auto")
        
        prompt_dir = os.path.join(get_app_dir(), 'prompts')
        if os.path.exists(prompt_dir):
            for file in os.listdir(prompt_dir):
                if file.startswith("prompt_") and file.endswith(".txt"):
                    subj_id = file.replace("prompt_", "").replace(".txt", "")
                    label = subj_id.capitalize()
                    with open(os.path.join(prompt_dir, file), 'r', encoding='utf-8') as f:
                        content = f.read()
                        match = re.search(r'\[SUBJECT_LABEL\]\n(.+)', content)
                        if match:
                            label = match.group(1).strip()
                    self.cb_subject.addItem(f"{label} ({subj_id})", subj_id)

    def open_add_subject_dialog(self):
        dialog = AddSubjectDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_subjects()

    def browse_original(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file đề gốc", "", "Word/PDF (*.docx *.pdf)")
        if path:
            self.lbl_orig.setText(path)

    def browse_answer(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file đáp án Excel", "", "Excel (*.xlsx)")
        if path:
            self.lbl_ans.setText(path)

    def browse_shuffled(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Chọn các file đề trộn", "", "Word/PDF (*.docx *.pdf)")
        for p in paths:
            if p not in self.shuffled_files_list:
                self.shuffled_files_list.append(p)
                self.list_shuffled.addItem(os.path.basename(p))

    def open_prompt_manager(self):
        dialog = PromptManagerDialog(self)
        dialog.exec_()
        # Load lại danh sách môn học ở màn hình chính phòng trường hợp có file mới
        self.load_subjects()
    def clear_shuffled(self):
        self.shuffled_files_list.clear()
        self.list_shuffled.clear()

    def normal_output_written(self, text):
        cursor = self.txt_log.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(text)
        self.txt_log.setTextCursor(cursor)
        self.txt_log.ensureCursorVisible()

    def start_checking(self):
        orig = self.lbl_orig.text().strip()
        ans = self.lbl_ans.text().strip()
        
        if not orig or not os.path.exists(orig):
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn file đề gốc hợp lệ.")
            return
        if not ans or not os.path.exists(ans):
            QMessageBox.warning(self, "Lỗi", "Vui lòng chọn file đáp án hợp lệ.")
            return
        if not self.shuffled_files_list:
            QMessageBox.warning(self, "Lỗi", "Vui lòng thêm ít nhất 1 file đề trộn.")
            return

        subject = self.cb_subject.currentData()
        vision_mode = self.chk_vision.isChecked()
        
        auto_pdf = False
        if subject in ['math', 'toan', 'vatly', 'hoahoc', 'sinhhoc']: 
            auto_pdf = True
        elif subject != 'auto':
            prompt_path = os.path.join(get_app_dir(), 'prompts', f'prompt_{subject}.txt')
            if os.path.exists(prompt_path):
                try:
                    with open(prompt_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        if re.search(r'\[AUTO_PDF\]\nTrue', content, re.IGNORECASE):
                            auto_pdf = True
                except:
                    pass

        self.btn_run.setEnabled(False)
        self.txt_log.clear()
        self.txt_log.append("Đang khởi tạo tiến trình kiểm tra...\n")

        self.worker = CheckWorker(orig, ans, self.shuffled_files_list, subject, vision_mode, auto_pdf)
        self.worker.log_signal.connect(self.normal_output_written)
        self.worker.finished_signal.connect(self.on_check_finished)
        self.worker.error_signal.connect(self.on_check_error)
        self.worker.start()

    def on_check_finished(self, msg):
        self.btn_run.setEnabled(True)
        self.txt_log.append(f"\n{msg}")
        QMessageBox.information(self, "Hoàn tất", "Đã kiểm tra xong! Vui lòng kiểm tra file Excel Kết quả tại thư mục chứa tool.")

    def on_check_error(self, err_msg):
        self.btn_run.setEnabled(True)
        self.txt_log.append(f"\n{err_msg}")
        QMessageBox.critical(self, "Lỗi Tiến Trình", err_msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainApp()
    window.show()
    sys.exit(app.exec_())