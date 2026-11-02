# 🎓 Admin Tool Dnc
**Công Cụ Tự Động Định Dạng Báo Cáo Khoa Học & Đồ Án - Trường Đại Học Nam Cần Thơ (DNC)**

---

## 🌟 Giới Thiệu
**Admin Tool Dnc** là bộ công cụ tự động hóa định dạng văn bản báo cáo khoa học, tiểu luận và đồ án học phần dành riêng cho sinh viên **Trường Đại học Nam Cần Thơ (DNC)**, đặc biệt là sinh viên **Khoa Công nghệ Thông tin**.

Chỉ cần nhập thông tin sinh viên và **dán nội dung văn bản thô vào**, công cụ sẽ tự động chuẩn hóa toàn bộ cấu trúc bài viết và xuất ra đồng thời 3 định dạng: **`.docx`**, **`.doc`**, và **`.pdf`** chất lượng cao.

---

## 🚀 Các Tính Năng Nổi Bật

### 1. 📘 Báo Cáo Cuối Kỳ (Chuẩn Quyết Định 309/QĐ-ĐHNCT)
- **Trang bìa chính & Trang phụ bìa**: Tự động chèn **Logo DNC chính thức**, khung viền mỹ thuật đôi (`twistedLines1`), tên đề tài cỡ chữ 20 in đậm, thông tin sinh viên và giảng viên hướng dẫn.
- **Tự động tạo MỤC LỤC**: Tự động nhận diện cấu trúc Chương, Mục 1.1, Mục 1.1.1, Tài liệu tham khảo, Phụ lục và tính toán số trang logic với hàng dấu chấm (`......`) sát lề phải.
- **Tự động tạo DANH SÁCH BẢNG & HÌNH**: Tự động lọc các tiêu đề `Bảng X.Y` và `Hình X.Y` để lập danh mục riêng.
- **Chuẩn căn lề & phông chữ**:
  - Phông chữ: **Times New Roman**, cỡ 13pt, dãn dòng 1.2, khoảng cách đoạn 6pt before.
  - Căn lề chuẩn: Trái **3.5 cm**, Phải **2.0 cm**, Trên **2.5 cm**, Dưới **2.5 cm**.
  - Đánh số trang: Số La Mã (`i, ii, iii...`) ở phần sơ khởi và số Ả Rập (`1, 2, 3...`) bắt đầu từ Chương 1.

### 2. 📗 Báo Cáo Giữa Kỳ (Chuẩn Quốc Tế Springer LNCS)
- **Bìa ngoài chuẩn DNC**: Thể hiện đủ 4 nội dung bắt buộc (Tên lớp, Tên môn, Tên giảng viên, Tên kỳ).
- **Cơ chế in 2 mặt (Duplex Printing)**: Tự động để trống mặt sau bìa vàng để phần bài báo khoa học nằm trọn vẹn trên giấy trắng A4.
- **Định dạng Springer LNCS**:
  - Căn lề chuẩn: Trái 4.4 cm, Phải 4.4 cm, Trên 4.5 cm, Dưới 4.5 cm (chiều rộng văn bản 12.2 cm).
  - Tiêu đề trang (Running headers) 2 mặt: Trang chẵn hiển thị Tên tác giả, Trang lẻ hiển thị Tên bài viết rút gọn.
  - Cấu trúc bài báo: Abstract, Keywords, các Section, Hình ảnh phân giải cao (300 DPI), và Danh mục tài liệu tham khảo đánh số `[1] - [17]`.

---

## 🛠️ Hướng Dẫn Cài Đặt & Sử Dụng

### Bước 1: Tải mã nguồn về máy
- **Cách 1**: Bấm vào nút xanh **Code** trên trang GitHub này $\rightarrow$ Chọn **Download ZIP** $\rightarrow$ Giải nén thư mục.
- **Cách 2**: Sử dụng lệnh Git Clone:
  ```bash
  git clone https://github.com/yamate2906/admin_tool_dnc.git
  cd admin_tool_dnc
  ```

### Bước 2: Cài đặt thư viện cần thiết
Mở Terminal / Command Prompt tại thư mục dự án và chạy lệnh:
```bash
pip install -r requirements.txt
```

### Bước 3: Khởi động phần mềm
Chỉ cần **nhấp đúp chuột vào file**:
👉 **`Chay_Tool_Dinh_Dang.bat`**

*(Hoặc chạy lệnh: `python app_gui.py`)*

### Bước 4: Sử dụng
1. Điền thông tin cá nhân (Họ tên, MSSV, Lớp, Giảng viên, Đề tài...).
2. Dán nội dung bài viết thô vào ô soạn thảo.
3. Chọn định dạng **Báo cáo Cuối kỳ (QĐ 309)** hoặc **Báo cáo Giữa kỳ (Springer LNCS)**.
4. Bấm nút **`🚀 TỰ ĐỘNG ĐỊNH DẠNG VÀ XUẤT TẤT CẢ FILE`**.
5. Mở thư mục để nhận ngay 3 file hoàn thiện: `.docx`, `.doc`, và `.pdf`!

---

## 📂 Cấu Trúc Thư Mục

```text
├── Chay_Tool_Dinh_Dang.bat                 # File khởi động nhanh 1-click
├── app_gui.py                              # Giao diện đồ họa phần mềm (Tkinter GUI)
├── tool_formatter.py                       # Bộ engine tự động định dạng & tạo Mục lục
├── dnc_logo.png                            # Logo chính thức Trường ĐH Nam Cần Thơ
├── requirements.txt                        # Danh sách thư viện Python phụ thuộc
└── README.md                               # Hướng dẫn chi tiết
```

---

## 👨‍💻 Tác Giả & Bản Quyền
- **Admin Đinh Văn Lắm** 


