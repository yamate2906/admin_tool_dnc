import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
import shutil
import os

def create_dnc_309_document():
    # Load Vietnamese source text
    source_doc = docx.Document('Random Forest và nền tảng WebGIS trong dự báo lưu lượng giao thông.docx')
    raw_paras = [p.text.strip() for p in source_doc.paragraphs if p.text.strip()]
    
    doc = docx.Document()
    
    # ---------------------------------------------------------------------------
    # COMMON STYLES AND HELPERS
    # ---------------------------------------------------------------------------
    def set_font(run, name='Times New Roman', size=Pt(13), bold=False, italic=False, color=None):
        run.font.name = name
        run.font.size = size
        run.bold = bold
        run.italic = italic
        if color:
            run.font.color.rgb = color

    def add_p(container, text="", font_name='Times New Roman', font_size=Pt(13),
              bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              space_before=Pt(6), space_after=Pt(0), line_spacing=1.2,
              first_line_indent=Cm(1.0), left_indent=Cm(0), right_indent=Cm(0)):
        p = container.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = space_before
        pf.space_after = space_after
        pf.line_spacing = line_spacing
        pf.first_line_indent = first_line_indent
        pf.left_indent = left_indent
        pf.right_indent = right_indent
        if text:
            r = p.add_run(text)
            set_font(r, font_name, font_size, bold, italic)
        return p

    def add_field(run, field_name):
        fldSimple = OxmlElement('w:fldSimple')
        fldSimple.set(qn('w:instr'), field_name)
        run._r.append(fldSimple)

    def add_decorative_border(section):
        sectPr = section._sectPr
        pgBorders = OxmlElement('w:pgBorders')
        pgBorders.set(qn('w:display'), 'firstPage')
        for side in ['top', 'left', 'bottom', 'right']:
            b = OxmlElement(f'w:{side}')
            b.set(qn('w:val'), 'twistedLines1')
            b.set(qn('w:sz'), '18')
            b.set(qn('w:space'), '4' if side == 'left' else '2')
            b.set(qn('w:color'), 'auto')
            pgBorders.append(b)
        sectPr.append(pgBorders)

    def remove_pg_borders(section):
        sectPr = section._sectPr
        pgBorders = sectPr.find(qn('w:pgBorders'))
        if pgBorders is not None:
            sectPr.remove(pgBorders)

    def set_page_margins(section):
        # Canh lề theo QĐ 309: Lề trái: 3,5 cm; Lề phải: 2,0 cm; Lề trên: 2,5 cm; Lề dưới: 2,5 cm
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(3.5)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.header_distance = Cm(1.0)
        section.footer_distance = Cm(1.0)

    def add_figure_with_image(image_path, num_str, title, source=None, width_cm=14.0):
        if os.path.exists(image_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(10)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.paragraph_format.first_line_indent = Cm(0)
            r_img = p_img.add_run()
            r_img.add_picture(image_path, width=Cm(width_cm))
        
        # Caption below figure, centered, font 13pt
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(4)
        p_cap.paragraph_format.space_after = Pt(2) if source else Pt(8)
        p_cap.paragraph_format.line_spacing = 1.15
        p_cap.paragraph_format.first_line_indent = Cm(0)
        r_cap = p_cap.add_run(f"Hình {num_str}. {title}")
        set_font(r_cap, 'Times New Roman', Pt(13), bold=True)
        
        if source:
            p_src = doc.add_paragraph()
            p_src.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_src.paragraph_format.space_before = Pt(0)
            p_src.paragraph_format.space_after = Pt(8)
            p_src.paragraph_format.first_line_indent = Cm(0)
            r_src = p_src.add_run(f"(Nguồn: {source})")
            set_font(r_src, 'Times New Roman', Pt(11), italic=True)

    # ---------------------------------------------------------------------------
    # SECTION 1: CÁC TRANG BÌA (Cover & Sub-Cover chuẩn QĐ 309 & Khoa CNTT)
    # ---------------------------------------------------------------------------
    sec_cover = doc.sections[0]
    set_page_margins(sec_cover)
    add_decorative_border(sec_cover)
    sec_cover.different_first_page_header_footer = True
    
    # TRANG BÌA CHÍNH (Trang bìa ngoài theo hình ảnh thực tế và mẫu khoa CNTT)
    add_p(doc, "TRƯỜNG ĐẠI HỌC NAM CẦN THƠ", font_size=Pt(14), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(6), space_after=Pt(3), line_spacing=1.1, first_line_indent=Cm(0))
    add_p(doc, "KHOA CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(10), line_spacing=1.1, first_line_indent=Cm(0))

    # Logo DNC chính thức
    if os.path.exists("dnc_logo.png"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(14)
        p_logo.paragraph_format.first_line_indent = Cm(0)
        r_logo = p_logo.add_run()
        r_logo.add_picture("dnc_logo.png", width=Cm(4.2))

    # Khối Tác giả sinh viên
    add_p(doc, "ĐINH VĂN LẮM", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "MSSV: 235292", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "LỚP: DH23TIN06", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(16), first_line_indent=Cm(0))

    # Khối Tên đề tài (Chuẩn cỡ 20 in đậm theo QĐ 309)
    add_p(doc, "ĐỀ TÀI", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(6), first_line_indent=Cm(0))
    add_p(doc, "XÂY DỰNG HỆ THỐNG WEBGIS DỰ ĐOÁN GIAO THÔNG\nBẰNG RANDOM FOREST",
          font_size=Pt(20), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(5), line_spacing=1.2, first_line_indent=Cm(0))
    add_p(doc, "(AN OVERVIEW STUDY ON THE APPLICATION OF THE RANDOM FOREST ALGORITHM\nAND WEBGIS PLATFORM IN TRAFFIC FORECASTING)",
          font_size=Pt(11.5), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(18), line_spacing=1.15, first_line_indent=Cm(0))

    # Khối Báo cáo học phần & Giảng viên hướng dẫn
    add_p(doc, "BÁO CÁO CUỐI KỲ", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "Phương Pháp Nghiên Cứu và Viết Báo Cáo Khoa Học", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "Học kỳ I", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "GIẢNG VIÊN HƯỚNG DẪN: TS. NGÔ HỒ ANH KHÔI", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "NGÀNH: CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "Mã số ngành: 7480201", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(26), first_line_indent=Cm(0))

    # Năm / Địa điểm
    add_p(doc, "Cần Thơ, Tháng 11/2026", font_size=Pt(13), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(0), first_line_indent=Cm(0))

    # TRANG PHỤ BÌA (Ngắt trang trong section theo QĐ 309 Mục 2.2)
    doc.add_page_break()
    add_p(doc, "TRƯỜNG ĐẠI HỌC NAM CẦN THƠ", font_size=Pt(14), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(12), space_after=Pt(3), line_spacing=1.1, first_line_indent=Cm(0))
    add_p(doc, "KHOA CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(10), line_spacing=1.1, first_line_indent=Cm(0))

    if os.path.exists("dnc_logo.png"):
        p_logo2 = doc.add_paragraph()
        p_logo2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo2.paragraph_format.space_before = Pt(4)
        p_logo2.paragraph_format.space_after = Pt(14)
        p_logo2.paragraph_format.first_line_indent = Cm(0)
        r_logo2 = p_logo2.add_run()
        r_logo2.add_picture("dnc_logo.png", width=Cm(4.2))

    add_p(doc, "ĐINH VĂN LẮM", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "MSSV: 235292", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "LỚP: DH23TIN06", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(16), first_line_indent=Cm(0))

    add_p(doc, "ĐỀ TÀI", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(6), first_line_indent=Cm(0))
    add_p(doc, "XÂY DỰNG HỆ THỐNG WEBGIS DỰ ĐOÁN GIAO THÔNG\nBẰNG RANDOM FOREST",
          font_size=Pt(20), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(5), line_spacing=1.2, first_line_indent=Cm(0))
    add_p(doc, "(AN OVERVIEW STUDY ON THE APPLICATION OF THE RANDOM FOREST ALGORITHM\nAND WEBGIS PLATFORM IN TRAFFIC FORECASTING)",
          font_size=Pt(11.5), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(18), line_spacing=1.15, first_line_indent=Cm(0))

    add_p(doc, "BÁO CÁO CUỐI KỲ", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "Học phần: Phương pháp nghiên cứu và viết báo cáo khoa học (Học kỳ I)", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "NGÀNH: CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "Mã số ngành: 7480201", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(10), first_line_indent=Cm(0))

    add_p(doc, "GIẢNG VIÊN HƯỚNG DẪN", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, "TS. NGÔ HỒ ANH KHÔI", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(20), first_line_indent=Cm(0))

    add_p(doc, "Cần Thơ, Tháng 11/2026", font_size=Pt(13), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(0), first_line_indent=Cm(0))

    # ---------------------------------------------------------------------------
    # SECTION 2: CÁC TRANG SƠ KHỞI (Front Matter: Roman numerals i, ii, iii...)
    # ĐÃ XÓA TRANG LỜI CẢM TẠ VÀ TRANG CAM KẾT THEO YÊU CẦU CỦA USER
    # ---------------------------------------------------------------------------
    sec_front = doc.add_section(WD_SECTION.NEW_PAGE)
    set_page_margins(sec_front)
    remove_pg_borders(sec_front) # Ensure NO border in Section 1!
    sec_front.footer.is_linked_to_previous = False
    sec_front.header.is_linked_to_previous = False
    
    # Configure Roman lower page numbers (i, ii, iii...) in footer
    footer_front_p = sec_front.footer.paragraphs[0]
    footer_front_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_roman = footer_front_p.add_run()
    set_font(r_roman, 'Times New Roman', Pt(12))
    add_field(r_roman, 'PAGE')
    
    sectPr_front = sec_front._sectPr
    pgNumType_front = OxmlElement('w:pgNumType')
    pgNumType_front.set(qn('w:fmt'), 'lowerRoman')
    pgNumType_front.set(qn('w:start'), '1')
    sectPr_front.append(pgNumType_front)

    # 2.6 MỤC LỤC
    add_p(doc, "MỤC LỤC", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(12), space_after=Pt(14), first_line_indent=Cm(0))
    
    def add_toc_line(title, page_str, level=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(3) if level==1 else Pt(1)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.15
        pf.tab_stops.add_tab_stop(Cm(15.5), WD_TAB_ALIGNMENT.RIGHT, docx.enum.text.WD_TAB_LEADER.DOTS)
        
        if level == 1:
            pf.left_indent = Cm(0)
            pf.first_line_indent = Cm(0)
            r1 = p.add_run(title)
            set_font(r1, 'Times New Roman', Pt(13), bold=True)
            r2 = p.add_run(f"\t{page_str}")
            set_font(r2, 'Times New Roman', Pt(13), bold=True)
        else:
            pf.left_indent = Cm(0.8)
            pf.first_line_indent = Cm(0)
            r1 = p.add_run(title)
            set_font(r1, 'Times New Roman', Pt(12.5), bold=False)
            r2 = p.add_run(f"\t{page_str}")
            set_font(r2, 'Times New Roman', Pt(12.5), bold=False)

    add_toc_line("DANH MỤC TỪ VIẾT TẮT", "ii", level=1)
    add_toc_line("DANH SÁCH BẢNG", "iii", level=1)
    add_toc_line("DANH SÁCH HÌNH", "iii", level=1)
    
    add_toc_line("CHƯƠNG 1: GIỚI THIỆU", "1", level=1)
    add_toc_line("1.1 Đặt vấn đề nghiên cứu", "1", level=2)
    add_toc_line("1.2 Mục tiêu nghiên cứu", "2", level=2)
    add_toc_line("1.3 Phạm vi và đối tượng nghiên cứu", "3", level=2)
    add_toc_line("1.4 Ý nghĩa khoa học và giá trị thực tiễn", "4", level=2)
    add_toc_line("1.5 Phương pháp nghiên cứu và bố cục báo cáo", "4", level=2)
    
    add_toc_line("CHƯƠNG 2: TỔNG QUAN NỀN TẢNG BẢN ĐỒ VÀ HỆ THỐNG THÔNG TIN ĐỊA LÝ (WEBGIS)", "6", level=1)
    add_toc_line("2.1 Khái niệm và vai trò của WebGIS trong quản lý giao thông", "6", level=2)
    add_toc_line("2.2 Đánh giá các nền tảng bản đồ số thương mại", "7", level=2)
    add_toc_line("2.3 Các giải pháp và thư viện bản đồ mã nguồn mở", "8", level=2)
    add_toc_line("2.4 Tiêu chí so sánh và kiến trúc phân tầng WebGIS", "10", level=2)
    
    add_toc_line("CHƯƠNG 3: MÃ NGUỒN VÀ THƯ VIỆN HỌC MÁY RANDOM FOREST", "13", level=1)
    add_toc_line("3.1 Nguyên lý thuật toán học tổ hợp Random Forest", "13", level=2)
    add_toc_line("3.2 Các siêu tham số cốt lõi của RandomForestRegressor", "14", level=2)
    add_toc_line("3.3 Thiết lập đường ống huấn luyện và tiền xử lý (Pipeline)", "16", level=2)
    add_toc_line("3.4 Chiến lược phân chia dữ liệu và giới hạn mô hình", "18", level=2)
    
    add_toc_line("CHƯƠNG 4: TẬP DỮ LIỆU VÀ ĐÁNH GIÁ MÔ HÌNH DỰ BÁO GIAO THÔNG", "20", level=1)
    add_toc_line("4.1 Tập dữ liệu chuẩn Metro Interstate Traffic Volume", "20", level=2)
    add_toc_line("4.2 Trích xuất đặc trưng không gian - thời gian", "21", level=2)
    add_toc_line("4.3 Phương pháp đo lường và các chỉ số đánh giá (RMSE, MAE, R²)", "23", level=2)
    add_toc_line("4.4 Lịch sử chủ đề và các khoảng trống nghiên cứu", "25", level=2)
    
    add_toc_line("CHƯƠNG 5: KẾT LUẬN VÀ KIẾN NGHỊ", "28", level=1)
    add_toc_line("5.1 Kết luận", "28", level=2)
    add_toc_line("5.2 Hạn chế của nghiên cứu", "29", level=2)
    add_toc_line("5.3 Kiến nghị và hướng phát triển tiếp theo", "29", level=2)
    
    add_toc_line("TÀI LIỆU THAM KHẢO", "31", level=1)
    add_toc_line("PHỤ LỤC", "33", level=1)

    # 2.9 DANH MỤC TỪ VIẾT TẮT
    doc.add_page_break()
    add_p(doc, "DANH MỤC TỪ VIẾT TẮT", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(12), space_after=Pt(14), first_line_indent=Cm(0))
    
    table_abbr = doc.add_table(rows=1, cols=3)
    table_abbr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_abbr.autofit = False
    
    hdr_cells = table_abbr.rows[0].cells
    hdr_titles = ["Từ viết tắt", "Tên tiếng Anh đầy đủ", "Ý nghĩa / Tên tiếng Việt"]
    widths = [Cm(2.5), Cm(6.5), Cm(6.5)]
    for cell, title, w in zip(hdr_cells, hdr_titles, widths):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(title)
        set_font(r, 'Times New Roman', Pt(12), bold=True)
        
    abbr_data = [
        ("GIS", "Geographic Information System", "Hệ thống thông tin địa lý"),
        ("WebGIS", "Web-based Geographic Information System", "Hệ thống thông tin địa lý trên nền tảng Web"),
        ("RF", "Random Forest", "Thuật toán Rừng ngẫu nhiên (học tổ hợp)"),
        ("ML", "Machine Learning", "Học máy / Trí tuệ nhân tạo"),
        ("OGC", "Open Geospatial Consortium", "Tổ chức tiêu chuẩn không gian địa lý mở"),
        ("WMS", "Web Map Service", "Dịch vụ hiển thị dữ liệu bản đồ dạng ảnh raster"),
        ("WFS", "Web Feature Service", "Dịch vụ truy xuất đối tượng dữ liệu hình học vector"),
        ("API", "Application Programming Interface", "Giao diện lập trình ứng dụng"),
        ("RMSE", "Root Mean Square Error", "Căn bậc hai của sai số toàn phương trung bình"),
        ("MAE", "Mean Absolute Error", "Sai số tuyệt đối trung bình"),
        ("MAPE", "Mean Absolute Percentage Error", "Sai số phần trăm tuyệt đối trung bình"),
        ("R²", "Coefficient of Determination", "Hệ số xác định (đo mức độ giải thích biến thiên)"),
        ("OOB", "Out-Of-Bag", "Dữ liệu ngoài túi mẫu trong Random Forest"),
        ("UCI", "University of California, Irvine", "Kho lưu trữ dữ liệu học máy chuẩn quốc tế"),
        ("ĐBSCL", "Mekong Delta", "Đồng bằng sông Cửu Long")
    ]
    
    for row_val in abbr_data:
        row_cells = table_abbr.add_row().cells
        for c_idx, (cell, val, w) in enumerate(zip(row_cells, row_val, widths)):
            cell.width = w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            r = p.add_run(val)
            set_font(r, 'Times New Roman', Pt(11.5), bold=(c_idx==0))

    # DANH SÁCH BẢNG VÀ DANH SÁCH HÌNH
    doc.add_page_break()
    add_p(doc, "DANH SÁCH BẢNG", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(12), space_after=Pt(10), first_line_indent=Cm(0))
    
    def add_list_item(item_title, page_str):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(2)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.15
        pf.tab_stops.add_tab_stop(Cm(15.5), WD_TAB_ALIGNMENT.RIGHT, docx.enum.text.WD_TAB_LEADER.DOTS)
        pf.left_indent = Cm(0)
        pf.first_line_indent = Cm(0)
        r1 = p.add_run(item_title)
        set_font(r1, 'Times New Roman', Pt(12.5), bold=False)
        r2 = p.add_run(f"\t{page_str}")
        set_font(r2, 'Times New Roman', Pt(12.5), bold=False)

    add_list_item("Bảng 2.1: So sánh đặc tính các nền tảng bản đồ số thương mại và mã nguồn mở", "9")
    add_list_item("Bảng 3.1: Các siêu tham số cốt lõi của bộ hồi quy RandomForestRegressor", "15")
    add_list_item("Bảng 4.1: Cấu trúc các trường thuộc tính trong tập dữ liệu Metro Interstate Traffic Volume", "20")
    add_list_item("Bảng 4.2: Các chỉ số sai số chính đánh giá hiệu năng mô hình dự báo", "23")

    add_p(doc, "DANH SÁCH HÌNH", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(18), space_after=Pt(10), first_line_indent=Cm(0))
    
    add_list_item("Hình 2.1: Sơ đồ kiến trúc phân tầng hệ thống WebGIS dự báo giao thông thông minh", "11")
    add_list_item("Hình 2.2: Luồng trao đổi dữ liệu không gian tuân theo các chuẩn dịch vụ mở OGC", "12")
    add_list_item("Hình 3.1: Cơ chế học tổ hợp biểu quyết và lấy mẫu ngẫu nhiên (Bootstrap) của Random Forest", "14")
    add_list_item("Hình 3.2: Quy trình đường ống (Pipeline) tiền xử lý và huấn luyện mô hình", "17")
    add_list_item("Hình 4.1: Biểu đồ mẫu hình phân bố lưu lượng phương tiện theo giờ và các ngày trong tuần", "22")
    add_list_item("Hình 4.2: Biểu đồ so sánh lưu lượng phương tiện thực tế và dự báo của mô hình Random Forest", "24")

    # ---------------------------------------------------------------------------
    # SECTION 3: PHẦN NỘI DUNG CHÍNH (Chương 1 đến Chương 5: Số Á-rập 1, 2, 3...)
    # ĐÃ XÓA KHUNG CỦA CHƯƠNG 1 (XÓA BỎ pgBorders KHỎI SECTION 3)
    # ---------------------------------------------------------------------------
    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    set_page_margins(sec_body)
    remove_pg_borders(sec_body) # Explicitly remove pgBorders so Chapter 1 has NO frame!
    sec_body.footer.is_linked_to_previous = False
    sec_body.header.is_linked_to_previous = False
    
    # Configure Arabic numbers (1, 2, 3...) in footer, centered
    footer_body_p = sec_body.footer.paragraphs[0]
    footer_body_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_arabic = footer_body_p.add_run()
    set_font(r_arabic, 'Times New Roman', Pt(12))
    add_field(r_arabic, 'PAGE')
    
    sectPr_body = sec_body._sectPr
    pgNumType_body = OxmlElement('w:pgNumType')
    pgNumType_body.set(qn('w:fmt'), 'decimal')
    pgNumType_body.set(qn('w:start'), '1')
    sectPr_body.append(pgNumType_body)

    def add_chapter_heading(ch_num, ch_title):
        add_p(doc, f"CHƯƠNG {ch_num}", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=Pt(18), space_after=Pt(4), first_line_indent=Cm(0))
        add_p(doc, ch_title.upper(), font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=Pt(0), space_after=Pt(16), line_spacing=1.2, first_line_indent=Cm(0))

    def add_heading_lvl1(num_str, title):
        add_p(doc, f"{num_str} {title.upper()}", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=Pt(12), space_after=Pt(4), first_line_indent=Cm(0))

    def add_heading_lvl2(num_str, title):
        add_p(doc, f"{num_str} {title}", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=Pt(8), space_after=Pt(3), left_indent=Cm(1.0), first_line_indent=Cm(0))

    def add_table_caption(num_str, title):
        add_p(doc, f"Bảng {num_str}. {title}", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=Pt(10), space_after=Pt(4), first_line_indent=Cm(0))

    # ===========================================================================
    # CHƯƠNG 1: GIỚI THIỆU (HOÀN TOÀN KHÔNG CÓ KHUNG VIỀN TRANG)
    # ===========================================================================
    add_chapter_heading(1, "GIỚI THIỆU")
    
    add_heading_lvl1("1.1", "Đặt vấn đề nghiên cứu")
    add_p(doc, raw_paras[7])
    add_p(doc, raw_paras[8])
    add_p(doc, "Tại Việt Nam, đặc biệt là các đô thị lớn và các trung tâm phát triển kinh tế vùng như thành phố Cần Thơ và khu vực Đồng bằng sông Cửu Long, áp lực giao thông ngày càng tăng cao đòi hỏi các cơ quan quản lý và người tham gia giao thông phải có công cụ giám sát trực quan và dự báo kịp thời. Một hệ thống tích hợp năng lực phân tích không gian của WebGIS kết hợp với thuật toán học máy Random Forest có khả năng mô hình hóa quan hệ phi tuyến giữa các yếu tố thời gian, thời tiết và mật độ phương tiện sẽ là giải pháp công nghệ mang lại giá trị thực tiễn to lớn.")

    add_heading_lvl1("1.2", "Mục tiêu nghiên cứu")
    add_heading_lvl2("1.2.1", "Mục tiêu tổng quát")
    add_p(doc, "Mục tiêu tổng quát của đề tài là nghiên cứu cơ sở lý thuyết, phương pháp luận khoa học và đề xuất kiến trúc hệ thống thông tin địa lý trên nền tảng Web (WebGIS) tích hợp thuật toán học máy Random Forest nhằm dự báo chính xác lưu lượng phương tiện giao thông, đồng thời trực quan hóa dữ liệu không gian hỗ trợ công tác điều hành và quản lý giao thông đô thị.")
    
    add_heading_lvl2("1.2.2", "Mục tiêu cụ thể")
    add_p(doc, "Để đạt được mục tiêu tổng quát trên, đề tài xác định các mục tiêu nghiên cứu cụ thể sau đây:")
    add_p(doc, "- Khảo sát, phân tích và so sánh các nền tảng bản đồ số thương mại (Google Maps, ArcGIS, Mapbox) và hệ sinh thái bản đồ mã nguồn mở (Leaflet, OpenLayers, MapLibre GL) theo các tiêu chuẩn kỹ thuật, chi phí bản quyền và khả năng vận hành độc lập.")
    add_p(doc, "- Nghiên cứu nguyên lý hoạt động, cấu trúc cây quyết định và các siêu tham số ảnh hưởng trực tiếp đến chất lượng dự báo của bộ ước lượng RandomForestRegressor trong thư viện scikit-learn.")
    add_p(doc, "- Xây dựng quy trình xử lý dữ liệu chuẩn hóa từ tập dữ liệu chuẩn Metro Interstate Traffic Volume, bao gồm kỹ thuật trích xuất đặc trưng chuỗi thời gian và mã hóa biến khí tượng.")
    add_p(doc, "- Thiết lập hệ thống chỉ số đo lường sai số (RMSE, MAE, R²) và chiến lược phân chia dữ liệu kiểm định theo trình tự thời gian (chronological split) nhằm phản ánh sát thực tế vận hành.")
    add_p(doc, "- Thiết kế kiến trúc giải pháp phân tầng cho ứng dụng WebGIS, phân tách tường minh giữa dịch vụ học máy dự báo và tầng trình diện bản đồ tương tác.")

    add_heading_lvl1("1.3", "Phạm vi và đối tượng nghiên cứu")
    add_heading_lvl2("1.3.1", "Phạm vi nghiên cứu")
    add_p(doc, raw_paras[10])
    add_p(doc, "Phạm vi nghiên cứu của đề tài tập trung vào bài toán hồi quy dự báo lưu lượng giao thông theo các bước thời gian xác định (ngắn hạn và trung hạn theo giờ), kết hợp dữ liệu thuộc tính thời gian và biến thời tiết khí tượng. Về mặt công nghệ, nghiên cứu giới hạn ở kiến trúc WebGIS phân tán với các tiêu chuẩn tương tác dữ liệu không gian OGC.")

    add_heading_lvl2("1.3.2", "Đối tượng nghiên cứu")
    add_p(doc, "Đối tượng nghiên cứu chính bao gồm: Thuật toán học máy Random Forest và các kỹ thuật tối ưu siêu tham số; Nền tảng bản đồ số WebGIS và các chuẩn giao tiếp dữ liệu địa lý; Dữ liệu lưu lượng phương tiện và chuỗi thời gian quan trắc giao thông.")

    add_heading_lvl1("1.4", "Ý nghĩa khoa học và thực tiễn của đề tài")
    add_p(doc, "Về mặt khoa học, đề tài góp phần hệ thống hóa cơ sở lý thuyết về việc kết hợp giữa công nghệ thông tin địa lý và học máy tổ hợp trong phân tích chuỗi thời gian không gian. Nghiên cứu chỉ rõ các giới hạn ngoại suy của mô hình dựa trên cây quyết định và đề xuất các nguyên tắc kiểm thử khách quan.")
    add_p(doc, "Về mặt thực tiễn, kết quả nghiên cứu cung cấp một tài liệu tham khảo chi tiết và định hướng kỹ thuật hoàn chỉnh cho các kỹ sư công nghệ thông tin, nhà quy hoạch đô thị và cơ quan quản lý giao thông trong việc xây dựng các hệ thống bản đồ giao thông thông minh với chi phí hợp lý và hiệu năng ổn định.")

    add_heading_lvl1("1.5", "Phương pháp nghiên cứu và bố cục báo cáo")
    add_p(doc, raw_paras[11])
    add_p(doc, raw_paras[9])

    # ===========================================================================
    # CHƯƠNG 2: TỔNG QUAN NỀN TẢNG BẢN ĐỒ VÀ HỆ THỐNG THÔNG TIN ĐỊA LÝ (WEBGIS)
    # ===========================================================================
    doc.add_page_break()
    add_chapter_heading(2, "TỔNG QUAN NỀN TẢNG BẢN ĐỒ VÀ HỆ THỐNG THÔNG TIN ĐỊA LÝ (WEBGIS)")
    
    add_heading_lvl1("2.1", "Khái niệm và vai trò của WebGIS trong quản lý giao thông")
    add_p(doc, "Hệ thống thông tin địa lý (GIS) là một hệ thống tích hợp phần cứng, phần mềm và dữ liệu nhằm thu thập, lưu trữ, thao tác, phân tích và trực quan hóa các dữ liệu mang thông tin vị trí không gian địa lý. Khi công nghệ web phát triển, WebGIS ra đời như một bước tiến tự nhiên, cho phép người dùng khai thác toàn bộ sức mạnh của GIS thông qua trình duyệt web mà không cần cài đặt phần mềm chuyên dụng nặng nề.")
    add_p(doc, "Trong lĩnh vực quản lý giao thông hiện đại, WebGIS đóng vai trò là một trung tâm điều phối và hiển thị dữ liệu trực quan theo thời gian thực. Bản đồ không chỉ hiển thị vị trí các tuyến đường mà còn đóng vai trò là lớp nền ngữ cảnh (spatial context) giúp người vận hành nhận biết tức thì mối tương quan giữa sự cố giao thông, thời tiết cực đoan và tình trạng ùn tắc tại các nút giao thông trọng điểm.")

    add_heading_lvl1("2.2", "Đánh giá các nền tảng bản đồ số thương mại")
    add_p(doc, raw_paras[13])
    add_p(doc, "Mỗi nhà cung cấp dịch vụ bản đồ thương mại đều sở hữu thế mạnh riêng biệt:")
    add_p(doc, "- Google Maps Platform: Sở hữu độ phủ dữ liệu mạng lưới đường bộ chi tiết nhất toàn cầu cùng hệ thống định vị địa chỉ xuất sắc. Tuy nhiên, mức phí sử dụng dựa trên lượt gọi API (pay-as-you-go) có thể tăng đột biến khi hệ thống tiếp cận lượng lớn người dùng công cộng.")
    add_p(doc, "- Esri ArcGIS Online: Cung cấp bộ công cụ phân tích không gian chuyên sâu và chuẩn mực công nghiệp cho các tổ chức nhà nước, nhưng đòi hỏi mức phí cấp phép bản quyền định kỳ rất cao và phụ thuộc sâu vào hệ sinh thái đóng của hãng.")
    add_p(doc, "- Mapbox và TomTom: Cung cấp khả năng tùy biến giao diện bản đồ linh hoạt thông qua vector tiles, định tuyến giao thông chuyên nghiệp, song chi phí duy trì cũng cần được tính toán kỹ lưỡng theo số lượng người dùng hoạt động hàng tháng (MAU).")

    add_heading_lvl1("2.3", "Các giải pháp và thư viện bản đồ mã nguồn mở")
    add_p(doc, raw_paras[14])
    add_p(doc, raw_paras[15])
    add_p(doc, raw_paras[16])
    
    # Bảng 2.1
    add_table_caption("2.1", "So sánh đặc tính các nền tảng bản đồ số thương mại và mã nguồn mở")
    t21 = doc.add_table(rows=1, cols=5)
    t21.alignment = WD_TABLE_ALIGNMENT.CENTER
    t21.autofit = False
    
    hdr21 = t21.rows[0].cells
    cols_t21 = ["Tiêu chí đánh giá", "Google Maps", "ArcGIS Online", "Mapbox", "Leaflet / MapLibre (Mã nguồn mở)"]
    w21 = [Cm(3.0), Cm(3.0), Cm(3.0), Cm(3.0), Cm(3.5)]
    for cell, text_h, w in zip(hdr21, cols_t21, w21):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text_h)
        set_font(r, 'Times New Roman', Pt(11), bold=True)
        
    data21 = [
        ("Mô hình chi phí", "Tính theo lượt gọi API (Pay-as-you-go)", "Thuê bao người dùng/năm rất cao", "Theo lượt tải bản đồ / MAU", "Miễn phí giấy phép mã nguồn mở"),
        ("Khả năng tự chủ máy chủ", "Hoàn toàn phụ thuộc hạ tầng Google", "Lưu trữ đám mây của Esri", "Đám mây Mapbox hoặc tự dựng tile", "Hoàn toàn tự chủ trên hạ tầng riêng"),
        ("Tùy biến phong cách bản đồ", "Hạn chế trong khuôn khổ cho phép", "Phong phú nhưng phức tạp", "Rất linh hoạt với Mapbox Studio", "Không giới hạn thông qua CSS/GL JS"),
        ("Tích hợp phân tích ML", "Chỉ hiển thị kết quả qua Layer", "Tích hợp công cụ GeoAI chuyên biệt", "Yêu cầu kết nối API bên ngoài", "Dễ dàng tích hợp với Python/FastAPI"),
        ("Bảo mật dữ liệu nội bộ", "Dữ liệu lưu chuyển qua máy chủ ngoài", "Theo chuẩn quản trị của Esri", "Cần ký kết thỏa thuận dịch vụ", "Tuyệt đối an toàn trong mạng nội bộ")
    ]
    for row_val in data21:
        row_cells = t21.add_row().cells
        for c_idx, (cell, val, w) in enumerate(zip(row_cells, row_val, w21)):
            cell.width = w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            r = p.add_run(val)
            set_font(r, 'Times New Roman', Pt(10.5), bold=(c_idx==0))
            
    add_p(doc, "(Nguồn: Tổng hợp từ tài liệu kỹ thuật và chính sách giá năm 2026)", font_size=Pt(11), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(2), space_after=Pt(8), first_line_indent=Cm(0))

    add_heading_lvl1("2.4", "Tiêu chí so sánh và kiến trúc phân tầng WebGIS")
    add_p(doc, raw_paras[17])
    add_p(doc, raw_paras[18])
    add_p(doc, raw_paras[19])
    
    # CHÈN HÌNH 2.1 THỰC TẾ
    add_figure_with_image('figures/fig_2_1_architecture.png', "2.1", 
                          "Sơ đồ kiến trúc phân tầng hệ thống WebGIS dự báo giao thông thông minh", 
                          source="Tác giả thiết kế và tổng hợp, 2026")
    
    add_p(doc, "Kiến trúc hệ thống đề xuất bao gồm 3 tầng chính:")
    add_p(doc, "1. Tầng dữ liệu và máy chủ phân tích (Backend Data & Analytics Layer): Chịu trách nhiệm lưu trữ dữ liệu lịch sử quan trắc, tiếp nhận dữ liệu thời tiết thời gian thực và thực thi mô hình học máy Random Forest thông qua API dịch vụ (RESTful API).")
    add_p(doc, "2. Tầng dịch vụ bản đồ (Geospatial Server Layer): Sử dụng máy chủ bản đồ mã nguồn mở (GeoServer hoặc Tile Server) để phát hành các lớp dữ liệu không gian theo các tiêu chuẩn OGC WMS/WFS hoặc Vector Tiles.")
    add_p(doc, "3. Tầng giao diện người dùng (Frontend Presentation Layer): Được xây dựng trên nền tảng thư viện bản đồ tương tác (như Leaflet hoặc MapLibre GL) kết hợp giao diện web hiện đại, giúp hiển thị trực quan các đoạn đường được mã hóa màu sắc theo mức độ lưu lượng dự báo.")

    # CHÈN HÌNH 2.2 THỰC TẾ
    add_figure_with_image('figures/fig_2_2_ogc_flow.png', "2.2", 
                          "Luồng trao đổi dữ liệu không gian tuân theo các chuẩn dịch vụ mở OGC", 
                          source="Tác giả tổng hợp theo chuẩn OGC, 2026")

    # ===========================================================================
    # CHƯƠNG 3: MÃ NGUỒN VÀ THƯ VIỆN HỌC MÁY RANDOM FOREST
    # ===========================================================================
    doc.add_page_break()
    add_chapter_heading(3, "MÃ NGUỒN VÀ THƯ VIỆN HỌC MÁY RANDOM FOREST")
    
    add_heading_lvl1("3.1", "Nguyên lý thuật toán học tổ hợp Random Forest")
    add_p(doc, raw_paras[21])
    add_p(doc, raw_paras[22])
    add_p(doc, raw_paras[23])
    
    # CHÈN HÌNH 3.1 THỰC TẾ
    add_figure_with_image('figures/fig_3_1_random_forest.png', "3.1", 
                          "Cơ chế học tổ hợp biểu quyết và lấy mẫu ngẫu nhiên (Bootstrap) của Random Forest", 
                          source="Mô phỏng theo Breiman, 2001")
    
    add_p(doc, "Về mặt toán học, nếu một mô hình rừng ngẫu nhiên bao gồm B cây quyết định hồi quy độc lập T_1(x), T_2(x), ..., T_B(x), dự báo đầu ra cho một vector đặc trưng x được xác định bằng công thức trung bình cộng:")
    add_p(doc, "f_RF(x) = (1 / B) * Σ_{b=1}^{B} T_b(x)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, first_line_indent=Cm(0))
    add_p(doc, "Nhờ cơ chế lấy mẫu có hoàn lại (bootstrap aggregating hay bagging) kết hợp với việc chọn ngẫu nhiên một tập con đặc trưng tại mỗi điểm phân nhánh, phương sai (variance) của mô hình tổng thể giảm đi đáng kể so với một cây quyết định đơn lẻ mà không làm tăng độ lệch (bias).")

    add_heading_lvl1("3.2", "Các siêu tham số cốt lõi của RandomForestRegressor")
    add_p(doc, raw_paras[24])
    
    # Bảng 3.1
    add_table_caption("3.1", "Các siêu tham số cốt lõi của bộ hồi quy RandomForestRegressor")
    t31 = doc.add_table(rows=1, cols=4)
    t31.alignment = WD_TABLE_ALIGNMENT.CENTER
    t31.autofit = False
    
    hdr31 = t31.rows[0].cells
    cols_t31 = ["Tên siêu tham số", "Giá trị mặc định", "Dải tìm kiếm khuyến nghị", "Ý nghĩa và tác động kỹ thuật"]
    w31 = [Cm(3.5), Cm(2.5), Cm(3.5), Cm(6.0)]
    for cell, text_h, w in zip(hdr31, cols_t31, w31):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text_h)
        set_font(r, 'Times New Roman', Pt(11), bold=True)
        
    data31 = [
        ("n_estimators", "100", "100 – 500", "Số lượng cây quyết định trong rừng. Số cây càng lớn mô hình càng ổn định nhưng chi phí tính toán tăng tuyến tính."),
        ("max_depth", "None", "10 – 30", "Độ sâu tối đa của mỗi cây. Cần giới hạn để ngăn cây phân nhánh quá sâu vào nhiễu dữ liệu huấn luyện."),
        ("min_samples_split", "2", "2 – 10", "Số lượng mẫu tối thiểu cần thiết để tiếp tục phân chia một nút nội bộ."),
        ("min_samples_leaf", "1", "1 – 5", "Số lượng mẫu tối thiểu bắt buộc tại mỗi nút lá. Tăng tham số này giúp làm mượt dự báo."),
        ("max_features", "1.0", "1.0, 'sqrt', 'log2'", "Số lượng đặc trưng được xem xét ngẫu nhiên khi tìm kiếm điểm chia nhánh tốt nhất."),
        ("n_jobs", "None", "-1", "Số luồng CPU được sử dụng song song trong quá trình huấn luyện và dự báo.")
    ]
    for row_val in data31:
        row_cells = t31.add_row().cells
        for c_idx, (cell, val, w) in enumerate(zip(row_cells, row_val, w31)):
            cell.width = w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            r = p.add_run(val)
            set_font(r, 'Times New Roman', Pt(10.5), bold=(c_idx==0))

    add_p(doc, "(Nguồn: Trích dẫn tài liệu scikit-learn và Probst et al., 2019)", font_size=Pt(11), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(2), space_after=Pt(8), first_line_indent=Cm(0))

    add_heading_lvl1("3.3", "Thiết lập đường ống huấn luyện và tiền xử lý (Pipeline)")
    add_p(doc, raw_paras[25])
    
    # CHÈN HÌNH 3.2 THỰC TẾ
    add_figure_with_image('figures/fig_3_2_pipeline.png', "3.2", 
                          "Quy trình đường ống (Pipeline) tiền xử lý và huấn luyện mô hình", 
                          source="Tác giả thiết kế theo quy chuẩn Scikit-Learn, 2026")
    
    add_p(doc, raw_paras[26])
    add_p(doc, raw_paras[27])

    add_heading_lvl1("3.4", "Chiến lược phân chia dữ liệu và giới hạn mô hình")
    add_p(doc, raw_paras[28])
    add_p(doc, "Hạn chế cốt lõi của thuật toán Random Forest trong bài toán dự báo chuỗi thời gian là tính chất không thể ngoại suy ngoài miền giá trị mục tiêu đã quan sát trong tập huấn luyện (inability to extrapolate). Khi lưu lượng giao thông tăng đột biến vượt ngưỡng cao nhất từng ghi nhận trong quá khứ (ví dụ trong các sự kiện di tản khẩn cấp hoặc lễ hội đặc biệt), mô hình chỉ có thể dự báo mức tối đa của các lá cây hiện có, không thể đưa ra con số cao hơn.")

    # ===========================================================================
    # CHƯƠNG 4: TẬP DỮ LIỆU VÀ ĐÁNH GIÁ MÔ HÌNH DỰ BÁO GIAO THÔNG
    # ===========================================================================
    doc.add_page_break()
    add_chapter_heading(4, "TẬP DỮ LIỆU VÀ ĐÁNH GIÁ MÔ HÌNH DỰ BÁO GIAO THÔNG")
    
    add_heading_lvl1("4.1", "Tập dữ liệu chuẩn Metro Interstate Traffic Volume")
    add_p(doc, raw_paras[30])
    add_p(doc, raw_paras[31])
    add_p(doc, raw_paras[32])
    
    # Bảng 4.1
    add_table_caption("4.1", "Cấu trúc các trường thuộc tính trong tập dữ liệu Metro Interstate Traffic Volume")
    t41 = doc.add_table(rows=1, cols=4)
    t41.alignment = WD_TABLE_ALIGNMENT.CENTER
    t41.autofit = False
    
    hdr41 = t41.rows[0].cells
    cols_t41 = ["Tên thuộc tính", "Kiểu dữ liệu", "Đơn vị đo lường", "Mô tả nội dung nghiệp vụ"]
    w41 = [Cm(3.5), Cm(2.5), Cm(3.0), Cm(6.5)]
    for cell, text_h, w in zip(hdr41, cols_t41, w41):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text_h)
        set_font(r, 'Times New Roman', Pt(11), bold=True)
        
    data41 = [
        ("holiday", "Chuỗi ký tự", "Phân loại", "Tên ngày nghỉ lễ quốc gia tại Hoa Kỳ (nếu có), hoặc 'None'."),
        ("temp", "Số thực (float)", "Độ Kelvin (K)", "Nhiệt độ không khí trung bình tại khu vực đo đạc."),
        ("rain_1h", "Số thực (float)", "Milimét (mm)", "Lượng mưa tích lũy ghi nhận trong khoảng thời gian một giờ."),
        ("snow_1h", "Số thực (float)", "Milimét (mm)", "Lượng tuyết rơi tích lũy ghi nhận trong một giờ."),
        ("clouds_all", "Số nguyên (int)", "Phần trăm (%)", "Tỷ lệ che phủ mây trên bầu trời (từ 0% đến 100%)."),
        ("weather_main", "Chuỗi ký tự", "Phân loại vĩ mô", "Nhóm điều kiện thời tiết khái quát (Clouds, Clear, Rain,...)."),
        ("weather_description", "Chuỗi ký tự", "Mô tả chi tiết", "Mô tả chi tiết khí tượng (light rain, sky is clear,...)."),
        ("date_time", "Dấu thời gian", "YYYY-MM-DD HH:mm:ss", "Mốc thời gian đo lường theo giờ địa phương."),
        ("traffic_volume", "Số nguyên (int)", "Số xe / giờ", "Biến mục tiêu: Tổng số lượng phương tiện đi qua trạm quan trắc.")
    ]
    for row_val in data41:
        row_cells = t41.add_row().cells
        for c_idx, (cell, val, w) in enumerate(zip(row_cells, row_val, w41)):
            cell.width = w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            r = p.add_run(val)
            set_font(r, 'Times New Roman', Pt(10.5), bold=(c_idx==0))

    add_p(doc, "(Nguồn: UCI Machine Learning Repository, Hogue 2019)", font_size=Pt(11), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(2), space_after=Pt(8), first_line_indent=Cm(0))

    add_heading_lvl1("4.2", "Trích xuất đặc trưng không gian - thời gian")
    add_p(doc, raw_paras[33])
    add_p(doc, raw_paras[34])
    add_p(doc, raw_paras[35])
    add_p(doc, raw_paras[36])
    add_p(doc, raw_paras[37])
    
    # CHÈN HÌNH 4.1 THỰC TẾ
    add_figure_with_image('figures/fig_4_1_traffic_pattern.png', "4.1", 
                          "Biểu đồ mẫu hình phân bố lưu lượng phương tiện theo giờ và các ngày trong tuần", 
                          source="Phân tích từ tập dữ liệu Metro Interstate, Hogue 2019")

    add_heading_lvl1("4.3", "Phương pháp đo lường và các chỉ số đánh giá (RMSE, MAE, R²)")
    add_p(doc, raw_paras[39])
    add_p(doc, raw_paras[40])
    add_p(doc, raw_paras[41])
    add_p(doc, raw_paras[42])
    add_p(doc, raw_paras[43])
    add_p(doc, raw_paras[44])
    add_p(doc, raw_paras[45])
    add_p(doc, raw_paras[46])

    # Bảng 4.2
    add_table_caption("4.2", "Các chỉ số sai số chính đánh giá hiệu năng mô hình dự báo")
    t42 = doc.add_table(rows=1, cols=4)
    t42.alignment = WD_TABLE_ALIGNMENT.CENTER
    t42.autofit = False
    
    hdr42 = t42.rows[0].cells
    cols_t42 = ["Chỉ số đánh giá", "Công thức toán học", "Đơn vị tính", "Đặc điểm và ý nghĩa đánh giá"]
    w42 = [Cm(2.5), Cm(5.0), Cm(2.5), Cm(5.5)]
    for cell, text_h, w in zip(hdr42, cols_t42, w42):
        cell.width = w
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(text_h)
        set_font(r, 'Times New Roman', Pt(11), bold=True)
        
    data42 = [
        ("MAE", "MAE = (1/n) * Σ |y_i - ŷ_i|", "Xe / giờ", "Đo sai số tuyệt đối trung bình, ít bị ảnh hưởng thái quá bởi ngoại lai."),
        ("RMSE", "RMSE = √[(1/n) * Σ (y_i - ŷ_i)²]", "Xe / giờ", "Phạt nặng các sai số lớn do phép bình phương, rất quan trọng khi phát hiện tắc nghẽn."),
        ("R²", "R² = 1 - [Σ(y_i - ŷ_i)² / Σ(y_i - ȳ)²]", "Không thứ nguyên (≤ 1.0)", "Hệ số xác định; đo lường tỷ lệ biến thiên của dữ liệu mà mô hình giải thích được.")
    ]
    for row_val in data42:
        row_cells = t42.add_row().cells
        for c_idx, (cell, val, w) in enumerate(zip(row_cells, row_val, w42)):
            cell.width = w
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.first_line_indent = Cm(0)
            r = p.add_run(val)
            set_font(r, 'Times New Roman', Pt(10.5), bold=(c_idx==0))

    add_p(doc, "(Nguồn: scikit-learn documentation và Pavlyuk, 2019)", font_size=Pt(11), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(2), space_after=Pt(8), first_line_indent=Cm(0))

    # CHÈN HÌNH 4.2 THỰC TẾ
    add_figure_with_image('figures/fig_4_2_prediction_comparison.png', "4.2", 
                          "Biểu đồ so sánh lưu lượng phương tiện thực tế và dự báo của mô hình Random Forest", 
                          source="Kết quả thực nghiệm trên tập kiểm thử 48 giờ liên tục, 2026")

    add_heading_lvl1("4.4", "Lịch sử chủ đề và các khoảng trống nghiên cứu")
    add_p(doc, raw_paras[48])
    add_p(doc, raw_paras[49])
    add_p(doc, raw_paras[50])
    add_p(doc, raw_paras[51])
    add_p(doc, raw_paras[52])
    add_p(doc, raw_paras[53])
    add_p(doc, raw_paras[54])
    add_p(doc, raw_paras[55])

    # ===========================================================================
    # CHƯƠNG 5: KẾT LUẬN VÀ KIẾN NGHỊ
    # ===========================================================================
    doc.add_page_break()
    add_chapter_heading(5, "KẾT LUẬN VÀ KIẾN NGHỊ")
    
    add_heading_lvl1("5.1", "Kết luận")
    add_p(doc, raw_paras[57])
    add_p(doc, raw_paras[58])
    add_p(doc, raw_paras[59])
    add_p(doc, raw_paras[60])
    add_p(doc, raw_paras[61])
    add_p(doc, raw_paras[62])

    add_heading_lvl1("5.2", "Hạn chế của nghiên cứu")
    add_p(doc, "Mặc dù đã đạt được những kết quả tổng quan toàn diện và xác định được mô hình kiến trúc khả thi, nghiên cứu vẫn còn một số điểm hạn chế cần được ghi nhận:")
    add_p(doc, "- Dữ liệu thực nghiệm phân tích trong đề tài mới chủ yếu dựa trên bộ dữ liệu quan trắc chuẩn Metro Interstate Traffic Volume tại một trạm quan trắc đơn lẻ ở Hoa Kỳ, chưa phản ánh đầy đủ đặc thù giao thông hỗn hợp với tỷ lệ xe hai bánh áp đảo như tại các đô thị Việt Nam.")
    add_p(doc, "- Mô hình Random Forest có độ phức tạp về bộ nhớ và kích thước tệp khi lưu trữ số lượng lớn cây quyết định, có thể tạo độ trễ nhất định khi triển khai trên các thiết bị biên hoặc máy chủ tài nguyên giới hạn.")
    add_p(doc, "- Chưa thực hiện kiểm thử tự động khả năng tự thích ứng của mô hình khi xảy ra hiện tượng trôi dạt dữ liệu (data drift) trong các giai đoạn biến động giao thông kéo dài.")

    add_heading_lvl1("5.3", "Kiến nghị và hướng phát triển tiếp theo")
    add_p(doc, raw_paras[63])
    add_p(doc, raw_paras[64])
    add_p(doc, "Từ những phát hiện và kết quả đạt được, đề tài kiến nghị một số phương hướng triển khai ứng dụng thực tế sau:")
    add_p(doc, "- Đối với nhà trường và Khoa Công nghệ Thông tin: Tiếp tục khuyến khích phát triển các đề tài nghiên cứu liên ngành giữa Trí tuệ nhân tạo (AI/Machine Learning) và Hệ thống thông tin địa lý (GIS), xây dựng các phòng thí nghiệm mô phỏng dữ liệu giao thông đô thị thông minh phục vụ học tập và nghiên cứu khoa học của sinh viên.")
    add_p(doc, "- Đối với các đơn vị vận hành giao thông địa phương: Từng bước thí điểm ứng dụng công nghệ WebGIS mã nguồn mở (như MapLibre GL, GeoServer) kết hợp với các mô hình học máy phân tích dữ liệu cảm biến camera để tối ưu hóa điều phối đèn tín hiệu và cung cấp thông tin cảnh báo ùn tắc cho người dân thành phố Cần Thơ.")

    # ===========================================================================
    # TÀI LIỆU THAM KHẢO (Theo đúng quy định chuẩn mục 2.11 của QĐ 309)
    # ===========================================================================
    doc.add_page_break()
    add_p(doc, "TÀI LIỆU THAM KHẢO", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(18), space_after=Pt(14), first_line_indent=Cm(0))
    
    references_vietnamese = [
        ("Atlas", "2026", "Is ArcGIS free? Atlas Blog, truy cập ngày 08/10/2026, từ https://atlas.co/blog/is-arcgis-free/"),
        ("Breiman, L.", "2001", "Random forests. Machine Learning, 45(1): 5–32. https://doi.org/10.1023/A:1010933404324"),
        ("CARTO", "2026", "CARTO Pricing and Enterprise Plans. CARTO Platform, truy cập ngày 08/10/2026, từ https://carto.com/pricing"),
        ("Esri", "2026", "Pricing for ArcGIS Online user types. Esri Official Portal, truy cập ngày 08/10/2026, từ https://www.esri.com/en-us/arcgis/products/arcgis-online/pricing/user-types"),
        ("Google", "2026", "Google Maps Platform pricing sheet. Google Cloud, truy cập ngày 08/10/2026, từ https://mapsplatform.google.com/pricing/"),
        ("Hogue, J.", "2019", "Metro Interstate Traffic Volume [Tập dữ liệu chuẩn]. UCI Machine Learning Repository. https://doi.org/10.24432/C56S3T"),
        ("Mapbox", "2026", "Mapbox Pricing & API Limits. Mapbox Documentation, truy cập ngày 08/10/2026, từ https://www.mapbox.com/pricing"),
        ("Open Geospatial Consortium", "2020", "OGC API – Features: Core standard specification. OGC Standards, truy cập ngày 08/10/2026, từ https://www.ogc.org/standards/ogcapi-features"),
        ("Pavlyuk, D.", "2019", "Feature selection and extraction in spatiotemporal traffic forecasting: A systematic literature review. European Transport Research Review, 11(1): 6. https://doi.org/10.1186/s12544-019-0345-3"),
        ("Pedregosa, F., G. Varoquaux, A. Gramfort, V. Michel, B. Thirion, O. Grisel, M. Blondel, P. Prettenhofer, R. Weiss, V. Dubourg, J. Vanderplas, A. Passos, D. Cournapeau, M. Brucher, M. Perrot and E. Duchesnay", "2011", "Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12: 2825–2830"),
        ("Probst, P., M. N. Wright and A.-L. Boulesteix", "2019", "Hyperparameters and tuning strategies for random forest. WIREs Data Mining and Knowledge Discovery, 9(3): e1301. https://doi.org/10.1002/widm.1301"),
        ("Reksoprodjo, M.", "2021", "metro-interstate-traffic: Traffic volume benchmark dataset. OpenML Repository, truy cập ngày 08/10/2026, từ https://api.openml.org/d/43582"),
        ("scikit-learn developers", "2024", "Metrics and scoring: Quantifying the quality of predictions. Scikit-Learn User Guide, truy cập ngày 08/10/2026, từ https://scikit-learn.org/stable/modules/model_evaluation.html"),
        ("scikit-learn developers", "2024", "sklearn.ensemble.RandomForestRegressor API reference. Scikit-Learn Documentation, truy cập ngày 08/10/2026, từ https://scikit-learn.org/stable/modules/generated/sklearn.ensemble.RandomForestRegressor.html"),
        ("StoreRocket", "2025", "Is Mapbox free? Mapbox pricing explained. StoreRocket Guide, truy cập ngày 08/10/2026, từ https://storerocket.io/learn/is-mapbox-free"),
        ("TomTom", "2026", "TomTom Maps and Traffic API Pricing. TomTom Developer Portal, truy cập ngày 08/10/2026, từ https://developer.tomtom.com/pricing"),
        ("Woosmap", "2026", "Mapbox pricing in 2026: A decision-maker's cost breakdown. Woosmap Tech Blog, truy cập ngày 08/10/2026, từ https://www.woosmap.com/blog/mapbox-pricing-explained")
    ]
    
    for author, yr, rest in references_vietnamese:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf_ref = p_ref.paragraph_format
        pf_ref.space_before = Pt(0)
        pf_ref.space_after = Pt(4)
        pf_ref.line_spacing = 1.0
        pf_ref.left_indent = Cm(1.0)
        pf_ref.first_line_indent = Cm(-1.0)
        
        r_ref = p_ref.add_run(f"-  {author} ({yr}), {rest}.")
        set_font(r_ref, 'Times New Roman', Pt(12))

    # ===========================================================================
    # PHỤ LỤC
    # ===========================================================================
    doc.add_page_break()
    add_p(doc, "PHỤ LỤC", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(18), space_after=Pt(14), first_line_indent=Cm(0))
    
    add_p(doc, "PHỤ LỤC 1: SƠ ĐỒ THIẾT KẾ CSDL VÀ DỊCH VỤ DỰ BÁO WEBGIS", font_size=Pt(13), bold=True,
          space_before=Pt(8), space_after=Pt(4), first_line_indent=Cm(0))
    add_p(doc, "Hệ thống WebGIS dự báo giao thông sử dụng kiến trúc dữ liệu không gian lưu trữ trên cơ sở dữ liệu quan hệ kết hợp thành phần không gian PostgreSQL/PostGIS. Các điểm nút và đoạn đường giao thông được lưu trữ dưới dạng bảng thuộc tính không gian:")
    add_p(doc, "- Bảng road_segments: Chứa thuộc tính hình học không gian (LineString) cùng các trường id, name, road_type, speed_limit, length_m.")
    add_p(doc, "- Bảng traffic_measurements: Lưu trữ chuỗi thời gian quan trắc lưu lượng gồm segment_id, timestamp, volume, weather_condition, temp_celsius.")
    add_p(doc, "- Bảng traffic_predictions: Lưu trữ kết quả dự báo của mô hình Random Forest gồm segment_id, predicted_timestamp, predicted_volume, confidence_interval.")

    add_p(doc, "PHỤ LỤC 2: DANH SÁCH THƯ VIỆN LẬP TRÌNH VÀ MÔI TRƯỜNG THỰC THI", font_size=Pt(13), bold=True,
          space_before=Pt(14), space_after=Pt(4), first_line_indent=Cm(0))
    add_p(doc, "1. Ngôn ngữ và môi trường máy học:")
    add_p(doc, "- Python phiên bản 3.10 trở lên.")
    add_p(doc, "- scikit-learn: Thư viện triển khai RandomForestRegressor, Pipeline, GridSearchCV.")
    add_p(doc, "- pandas & numpy: Thao tác dữ liệu bảng và tính toán ma trận số học.")
    add_p(doc, "- joblib: Tuần tự hóa và lưu trữ tệp mô hình đã huấn luyện (.joblib / .pkl).")
    add_p(doc, "2. Nền tảng dịch vụ web và bản đồ:")
    add_p(doc, "- FastAPI: Xây dựng RESTful API phục vụ các yêu cầu dự báo giao thông theo thời gian thực.")
    add_p(doc, "- Leaflet.js / MapLibre GL JS: Thư viện hiển thị bản đồ trực quan tương tác trên trình duyệt.")
    add_p(doc, "- GeoJSON: Định dạng dữ liệu không gian chuẩn trao đổi giữa backend FastAPI và bản đồ web frontend.")

    out_docx = "Bao_cao_Cuoi_ki_DNC_309_Dinh_Van_Lam.docx"
    doc.save(out_docx)
    print(f"Generated {out_docx} successfully.")
    
    out_doc = "Bao_cao_Cuoi_ki_DNC_309_Dinh_Van_Lam.doc"
    shutil.copyfile(out_docx, out_doc)
    print(f"Copied to {out_doc}.")

if __name__ == "__main__":
    create_dnc_309_document()
