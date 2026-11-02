import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import shutil
import threading
import sys
import tool_formatter

class AutoReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DNC AutoReport Formatter - Công Cụ Định Dạng Báo Cáo Tự Động ĐH Nam Cần Thơ")
        self.root.geometry("900x750")
        self.root.minsize(800, 600)
        self.lbl_status = None
        
        # Apply TTK Theme
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass

        # Main Layout (Bottom-anchored controls, Top-anchored tabs)
        self.create_header()
        self.create_log_area()
        self.create_action_bar()
        self.create_tabs()

    def create_header(self):
        header_frame = tk.Frame(self.root, bg="#0d47a1", padx=15, pady=10)
        header_frame.pack(fill=tk.X)
        
        title_label = tk.Label(header_frame, text="CÔNG CỤ TỰ ĐỘNG ĐỊNH DẠNG BÁO CÁO KHOA HỌC DNC",
                               font=("Arial", 14, "bold"), fg="#ffffff", bg="#0d47a1")
        title_label.pack(anchor="w")
        
        sub_label = tk.Label(header_frame, text="Hỗ trợ chuẩn QĐ 309/QĐ-ĐHNCT (Cuối kỳ) & Chuẩn Springer LNCS (Giữa kỳ) - Tự động xuất .docx, .doc, .pdf",
                             font=("Arial", 9), fg="#e3f2fd", bg="#0d47a1")
        sub_label.pack(anchor="w")

    def create_tabs(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=8)

        # Tab 1: Metadata
        self.tab_meta = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_meta, text=" 📋 1. Thông Tin Trang Bìa ")
        self.build_meta_tab()

        # Tab 2: Content
        self.tab_content = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_content, text=" 📝 2. Dán Nội Dung Báo Cáo ")
        self.build_content_tab()

    def build_meta_tab(self):
        # Grid layout for metadata inputs
        fields = [
            ("Họ và tên sinh viên:", "ho_ten", "ĐINH VĂN LẮM"),
            ("Mã số sinh viên (MSSV):", "mssv", "235292"),
            ("Lớp sinh hoạt:", "lop", "DH23TIN06"),
            ("Tên đề tài (Tiếng Việt):", "ten_de_tai", "XÂY DỰNG HỆ THỐNG WEBGIS DỰ ĐOÁN GIAO THÔNG BẰNG RANDOM FOREST"),
            ("Tên đề tài (Tiếng Anh):", "ten_de_tai_tieng_anh", "An Overview Study on the Application of the Random Forest Algorithm and WebGIS Platform in Traffic Forecasting"),
            ("Tên học phần / Môn học:", "mon_hoc", "Phương Pháp Nghiên Cứu và Viết Báo Cáo Khoa Học"),
            ("Học kỳ thực hiện:", "hoc_ky", "Học kỳ I"),
            ("Giảng viên hướng dẫn:", "giang_vien", "TS. NGÔ HỒ ANH KHÔI"),
            ("Khoa quản lý:", "khoa", "Khoa Công nghệ Thông tin"),
            ("Ngành đào tạo:", "nganh", "Công nghệ Thông tin"),
            ("Mã số ngành:", "ma_nganh", "7480201"),
            ("Địa điểm và Thời gian:", "ngay_thang", "Cần Thơ, Tháng 11/2026")
        ]
        
        self.entries = {}
        for idx, (label_text, key, default_val) in enumerate(fields):
            row = idx // 2
            col = (idx % 2) * 2
            
            lbl = ttk.Label(self.tab_meta, text=label_text, font=("Arial", 9, "bold"))
            lbl.grid(row=row*2, column=col, sticky="w", padx=6, pady=(6, 2))
            
            ent = ttk.Entry(self.tab_meta, font=("Arial", 9), width=45)
            ent.insert(0, default_val)
            ent.grid(row=row*2+1, column=col, sticky="ew", padx=6, pady=(0, 6))
            self.entries[key] = ent
            
        self.tab_meta.columnconfigure(0, weight=1)
        self.tab_meta.columnconfigure(2, weight=1)

    def build_content_tab(self):
        toolbar = ttk.Frame(self.tab_content)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        btn_paste = ttk.Button(toolbar, text="📋 Dán từ Clipboard", command=self.paste_from_clipboard)
        btn_paste.pack(side=tk.LEFT, padx=3)

        btn_open = ttk.Button(toolbar, text="📂 Nhập file văn bản (.txt)", command=self.open_text_file)
        btn_open.pack(side=tk.LEFT, padx=3)

        btn_sample = ttk.Button(toolbar, text="💡 Dán mẫu thử nghiệm", command=self.insert_sample_content)
        btn_sample.pack(side=tk.LEFT, padx=3)

        btn_clear = ttk.Button(toolbar, text="🗑 Xóa trắng", command=self.clear_content)
        btn_clear.pack(side=tk.RIGHT, padx=3)

        # Text input area with scrollbar
        text_container = ttk.Frame(self.tab_content)
        text_container.pack(fill=tk.BOTH, expand=True)

        self.text_editor = tk.Text(text_container, wrap=tk.WORD, font=("Times New Roman", 11),
                                   padx=10, pady=10, undo=True)
        scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, command=self.text_editor.yview)
        self.text_editor.configure(yscrollcommand=scrollbar.set)

        self.text_editor.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load sample content by default
        self.insert_sample_content()

    def create_action_bar(self):
        action_frame = ttk.LabelFrame(self.root, text=" ⚙️ Tùy Chọn Định Dạng & Xuất Báo Cáo ", padding=10)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 6))

        # Format choice
        self.format_var = tk.StringVar(value="DNC_309")
        rb1 = ttk.Radiobutton(action_frame, text="📘 Báo Cáo Cuối Kỳ (Chuẩn QĐ 309/ĐHNCT - Tự động tạo Mục Lục, Bảng, Hình)",
                              variable=self.format_var, value="DNC_309")
        rb1.grid(row=0, column=0, sticky="w", padx=10, pady=3)

        rb2 = ttk.Radiobutton(action_frame, text="📗 Báo Cáo Giữa Kỳ (Chuẩn Springer LNCS - Tiếng Anh)",
                              variable=self.format_var, value="SPRINGER")
        rb2.grid(row=0, column=1, sticky="w", padx=10, pady=3)

        # Output file name
        lbl_name = ttk.Label(action_frame, text="Tên file xuất ra:")
        lbl_name.grid(row=1, column=0, sticky="w", padx=10, pady=3)
        self.ent_filename = ttk.Entry(action_frame, width=40)
        self.ent_filename.insert(0, "Bao_cao_Xuat_Tu_Dong")
        self.ent_filename.grid(row=1, column=1, sticky="w", padx=10, pady=3)

        # PDF Export Checkbox
        self.export_pdf_var = tk.BooleanVar(value=True)
        chk_pdf = ttk.Checkbutton(action_frame, text="Tự động xuất cả file .PDF cùng lúc (chất lượng cao)",
                                  variable=self.export_pdf_var)
        chk_pdf.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=3)

        # Big Action Button
        self.btn_run = tk.Button(action_frame, text="🚀 TỰ ĐỘNG ĐỊNH DẠNG VÀ XUẤT TẤT CẢ FILE (.DOCX, .DOC, .PDF)",
                                 font=("Arial", 11, "bold"), bg="#1976d2", fg="#ffffff",
                                 activebackground="#1565c0", activeforeground="#ffffff",
                                 padx=15, pady=8, cursor="hand2", command=self.start_processing)
        self.btn_run.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=(8, 2))

    def create_log_area(self):
        log_frame = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        log_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.lbl_status = ttk.Label(log_frame, text="Trạng thái: Sẵn sàng thực hiện.", font=("Arial", 9, "italic"))
        self.lbl_status.pack(side=tk.LEFT)

        self.btn_open_folder = ttk.Button(log_frame, text="📁 Mở Thư Mục Chứa File", command=self.open_current_directory)
        self.btn_open_folder.pack(side=tk.RIGHT)

    def paste_from_clipboard(self):
        try:
            clipboard_text = self.root.clipboard_get()
            self.text_editor.insert(tk.INSERT, clipboard_text)
            self.lbl_status.config(text="Đã dán văn bản từ Clipboard thành công.")
        except Exception as e:
            messagebox.showwarning("Thông báo", "Không tìm thấy văn bản trong Clipboard!")

    def clear_content(self):
        self.text_editor.delete("1.0", tk.END)
        self.lbl_status.config(text="Đã xóa sạch nội dung soạn thảo.")

    def open_text_file(self):
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.text_editor.delete("1.0", tk.END)
                self.text_editor.insert("1.0", content)
                self.lbl_status.config(text=f"Đã mở file: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {e}")

    def insert_sample_content(self):
        sample = """CHƯƠNG 1: GIỚI THIỆU VẤN ĐỀ NGHIÊN CỨU
1.1 Đặt vấn đề nghiên cứu
Trong bối cảnh đô thị hóa ngày càng diễn ra nhanh chóng tại Việt Nam và trên thế giới, tình trạng ùn tắc giao thông đã và đang trở thành một thách thức lớn đối với công tác quy hoạch và vận hành hạ tầng kỹ thuật đô thị. Sự quá tải của các mạng lưới đường giao thông huyết mạch không chỉ gây tổn thất lớn về thời gian, kinh tế mà còn trực tiếp làm gia tăng lượng khí thải độc hại vào môi trường sống.
Để giải quyết bài toán giao thông phức tạp này, việc ứng dụng các thuật toán máy học hiện đại nhằm xây dựng hệ thống dự báo lưu lượng giao thông tự động theo thời gian thực đóng vai trò đặc biệt quan trọng. 

1.2 Mục tiêu nghiên cứu của đề tài
Đề tài hướng tới việc nghiên cứu và ứng dụng thuật toán học máy Random Forest kết hợp cùng nền tảng bản đồ số WebGIS để giải quyết bài toán dự báo lưu lượng giao thông:
- Xây dựng quy trình thu thập và tiền xử lý dữ liệu lưu lượng giao thông theo chuỗi thời gian.
- Huấn luyện và đánh giá mô hình Random Forest Regressor trên tập dữ liệu đo đạc thực tế.
- Thiết kế kiến trúc dịch vụ bản đồ số WebGIS phục vụ trực quan hóa dữ liệu dự báo tương tác cao.

CHƯƠNG 2: CƠ SỞ LÝ THUYẾT VÀ PHƯƠNG PHÁP NGHIÊN CỨU
2.1 Tổng quan về giải thuật Random Forest
Thuật toán Random Forest là một kỹ thuật học máy có giám sát thuộc nhóm Ensemble Learning, hoạt động dựa trên nguyên lý kết hợp dự báo từ một tập hợp rất nhiều cây quyết định độc lập (Decision Trees) thông qua cơ chế lấy mẫu có hoàn lại (Bootstrap Aggregating / Bagging).
Nhờ nguyên lý bỏ phiếu số đông và tính trung bình các cây con, Random Forest có khả năng hạn chế tối đa hiện tượng học vẹt (overfitting) và duy trì độ chính xác cao trên các tập dữ liệu không gian - thời gian phức tạp.

2.2 Kiến trúc tích hợp nền tảng WebGIS
Hệ thống WebGIS dự báo giao thông được thiết kế theo mô hình 3 tầng dịch vụ:
- Tầng cơ sở dữ liệu: Lưu trữ không gian địa lý dựa trên PostgreSQL kết hợp thành phần PostGIS.
- Tầng dịch vụ backend: Xây dựng bằng ngôn ngữ Python với framework FastAPI để xử lý dự báo theo thời gian thực.
- Tầng giao diện người dùng: Tích hợp thư viện Leaflet / MapLibre GL JS để hiển thị bản đồ trực quan."""
        self.text_editor.delete("1.0", tk.END)
        self.text_editor.insert("1.0", sample)
        if hasattr(self, 'lbl_status') and self.lbl_status:
            self.lbl_status.config(text="Đã nạp văn bản mẫu cấu trúc báo cáo.")

    def open_current_directory(self):
        os.startfile(os.path.abspath("."))

    def start_processing(self):
        content = self.text_editor.get("1.0", tk.END).strip()
        if not content:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập hoặc dán nội dung bài báo cáo trước khi xuất!")
            return

        self.btn_run.config(state=tk.DISABLED, text="⏳ Đang tự động định dạng và xuất file, vui lòng chờ...")
        self.lbl_status.config(text="Đang xử lý tạo trang bìa, chuẩn hóa phông chữ và căn lề...")

        # Run in thread to keep GUI responsive
        threading.Thread(target=self.run_formatting_worker, args=(content,), daemon=True).start()

    def run_formatting_worker(self, content):
        try:
            meta = {k: ent.get().strip() for k, ent in self.entries.items()}
            meta["logo_path"] = "dnc_logo.png"

            base_name = self.ent_filename.get().strip() or "Bao_cao_Xuat_Tu_Dong"
            docx_file = f"{base_name}.docx"
            doc_file = f"{base_name}.doc"
            pdf_file = f"{base_name}.pdf"

            is_springer = (self.format_var.get() == "SPRINGER")

            if is_springer:
                tool_formatter.parse_and_format_springer(content, meta, docx_file)
            else:
                tool_formatter.parse_and_format_dnc_309(content, meta, docx_file)

            # Copy to .doc
            shutil.copyfile(docx_file, doc_file)

            # Export to PDF if requested
            pdf_success = False
            if self.export_pdf_var.get():
                try:
                    tool_formatter.convert_docx_to_pdf_win32(docx_file, pdf_file)
                    pdf_success = True
                except Exception as pdf_err:
                    print("PDF Error:", pdf_err)

            self.root.after(0, self.on_processing_complete, docx_file, doc_file, pdf_file if pdf_success else None)
        except Exception as e:
            self.root.after(0, self.on_processing_error, str(e))

    def on_processing_complete(self, docx_file, doc_file, pdf_file):
        self.btn_run.config(state=tk.NORMAL, text="🚀 TỰ ĐỘNG ĐỊNH DẠNG VÀ XUẤT TẤT CẢ FILE (.DOCX, .DOC, .PDF)")
        self.lbl_status.config(text=f"✅ Hoàn thành xuất sắc! Đã xuất: {docx_file}, {doc_file}" + (f", {pdf_file}" if pdf_file else ""))

        msg = f"Đã định dạng và xuất báo cáo thành công!\n\nCác file đã được tạo tại thư mục làm việc:\n1. {docx_file} (Word hiện đại)\n2. {doc_file} (Word 97-2003)"
        if pdf_file:
            msg += f"\n3. {pdf_file} (File PDF chuẩn in ấn)"
        msg += "\n\nBạn có muốn mở ngay thư mục chứa file không?"

        if messagebox.askyesno("Thành công!", msg):
            self.open_current_directory()

    def on_processing_error(self, err_msg):
        self.btn_run.config(state=tk.NORMAL, text="🚀 TỰ ĐỘNG ĐỊNH DẠNG VÀ XUẤT TẤT CẢ FILE (.DOCX, .DOC, .PDF)")
        self.lbl_status.config(text=f"❌ Có lỗi xảy ra: {err_msg}")
        messagebox.showerror("Lỗi thực thi", f"Không thể hoàn thành định dạng:\n{err_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoReportApp(root)
    root.mainloop()

