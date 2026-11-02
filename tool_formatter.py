import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
import shutil
import os
import re
import sys

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

def set_dnc_margins(section):
    # Lề trái: 3.5 cm; Lề phải: 2.0 cm; Lề trên: 2.5 cm; Lề dưới: 2.5 cm
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.header_distance = Cm(1.0)
    section.footer_distance = Cm(1.0)

def set_springer_margins(section):
    # Chuẩn Springer LNCS: Lề trái: 4.4 cm; Lề phải: 4.4 cm; Lề trên: 4.5 cm; Lề dưới: 4.5 cm
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(4.4)
    section.right_margin = Cm(4.4)
    section.top_margin = Cm(4.5)
    section.bottom_margin = Cm(4.5)
    section.header_distance = Cm(3.2)
    section.footer_distance = Cm(3.2)

def add_dnc_cover(doc, meta, is_midterm=False):
    sec = doc.sections[0]
    set_dnc_margins(sec)
    add_decorative_border(sec)
    sec.different_first_page_header_footer = True
    
    # 1. Header trường & Khoa
    add_p(doc, "TRƯỜNG ĐẠI HỌC NAM CẦN THƠ", font_size=Pt(14), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(6), space_after=Pt(3), line_spacing=1.1, first_line_indent=Cm(0))
    add_p(doc, meta.get("khoa", "KHOA CÔNG NGHỆ THÔNG TIN").upper(), font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(10), line_spacing=1.1, first_line_indent=Cm(0))

    # 2. Logo DNC
    logo_path = meta.get("logo_path", "dnc_logo.png")
    if os.path.exists(logo_path):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(14)
        p_logo.paragraph_format.first_line_indent = Cm(0)
        r_logo = p_logo.add_run()
        r_logo.add_picture(logo_path, width=Cm(4.2))

    # 3. Thông tin sinh viên
    add_p(doc, meta.get("ho_ten", "ĐINH VĂN LẮM").upper(), font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, f"MSSV: {meta.get('mssv', '235292')}", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, f"LỚP: {meta.get('lop', 'DH23TIN06')}", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(16), first_line_indent=Cm(0))

    # 4. Tên đề tài (cỡ 20 in đậm theo QĐ 309)
    add_p(doc, "ĐỀ TÀI", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(6), first_line_indent=Cm(0))
    add_p(doc, meta.get("ten_de_tai", "XÂY DỰNG HỆ THỐNG WEBGIS DỰ ĐOÁN GIAO THÔNG\nBẰNG RANDOM FOREST").upper(),
          font_size=Pt(20), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(5), line_spacing=1.2, first_line_indent=Cm(0))
    
    subtitle = meta.get("ten_de_tai_tieng_anh", "")
    if subtitle:
        add_p(doc, f"({subtitle.upper()})", font_size=Pt(11.5), italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=Pt(0), space_after=Pt(18), line_spacing=1.15, first_line_indent=Cm(0))

    # 5. Khối Báo cáo học phần
    report_title = "BÁO CÁO GIỮA KỲ" if is_midterm else "BÁO CÁO CUỐI KỲ"
    add_p(doc, report_title, font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, meta.get("mon_hoc", "Phương Pháp Nghiên Cứu và Viết Báo Cáo Khoa Học"), font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, meta.get("hoc_ky", "Học kỳ I"), font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, f"GIẢNG VIÊN HƯỚNG DẪN: {meta.get('giang_vien', 'TS. NGÔ HỒ ANH KHÔI').upper()}", font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, f"NGÀNH: {meta.get('nganh', 'CÔNG NGHỆ THÔNG TIN').upper()}", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(3), first_line_indent=Cm(0))
    add_p(doc, f"Mã số ngành: {meta.get('ma_nganh', '7480201')}", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(26), first_line_indent=Cm(0))

    # 6. Địa điểm & Thời gian
    add_p(doc, meta.get("ngay_thang", "Cần Thơ, Tháng 11/2026"), font_size=Pt(13), bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(0), first_line_indent=Cm(0))

def extract_headings_and_pages(raw_text):
    lines = raw_text.split('\n')
    toc_items = []
    tables = []
    figures = []
    
    current_page = 1
    current_lines = 0
    is_first_chapter = True
    
    for line in lines:
        line_s = line.strip()
        if not line_s:
            continue
            
        # 1. Phát hiện Chương (Chương mới luôn ngắt trang mới)
        if re.match(r'^(CHƯƠNG|Chương)\s+\d+', line_s, re.IGNORECASE):
            if not is_first_chapter:
                current_page += 1
                current_lines = 0
            else:
                is_first_chapter = False
                current_page = 1
                current_lines = 0
            toc_items.append((line_s.upper(), current_page, 1))
            continue
            
        # 2. Phát hiện Tiêu đề lớn đặc biệt (TÀI LIỆU THAM KHẢO, PHỤ LỤC...)
        if re.match(r'^(TÀI LIỆU THAM KHẢO|PHỤ LỤC|KẾT LUẬN VÀ KIẾN NGHỊ)', line_s, re.IGNORECASE):
            current_page += 1
            current_lines = 0
            toc_items.append((line_s.upper(), current_page, 1))
            continue
            
        # 3. Phát hiện Mục cấp 1 (ví dụ: 1.1, 2.1)
        if re.match(r'^\d+\.\d+\s+[^\.]', line_s):
            toc_items.append((line_s, current_page, 2))
            current_lines += 2
            continue
            
        # 4. Phát hiện Mục cấp 2 (ví dụ: 1.1.1, 2.1.1)
        if re.match(r'^\d+\.\d+\.\d+\s+[^\.]', line_s):
            toc_items.append((line_s, current_page, 3))
            current_lines += 2
            continue
            
        # 5. Phát hiện Bảng
        if re.match(r'^(Bảng|Table)\s+\d+', line_s, re.IGNORECASE):
            tables.append((line_s, current_page))
            current_lines += 3
            continue
            
        # 6. Phát hiện Hình
        if re.match(r'^(Hình|Fig\.)\s+\d+', line_s, re.IGNORECASE):
            figures.append((line_s, current_page))
            current_lines += 3
            continue
            
        # Đoạn văn nội dung thông thường
        lines_est = max(1, len(line_s) // 80 + 1)
        current_lines += lines_est
        if current_lines >= 26:
            current_page += current_lines // 26
            current_lines = current_lines % 26
            
    return toc_items, tables, figures

def parse_and_format_dnc_309(raw_text, meta, output_docx):
    doc = docx.Document()
    
    # -----------------------------------------------------------------------
    # SECTION 0: TRANG BÌA CHÍNH VÀ TRANG PHỤ BÌA (Theo chuẩn QĐ 309)
    # -----------------------------------------------------------------------
    add_dnc_cover(doc, meta, is_midterm=False)
    
    # Trích xuất tự động mục lục, danh mục bảng, danh mục hình từ nội dung
    toc_items, tables, figures = extract_headings_and_pages(raw_text)
    
    # -----------------------------------------------------------------------
    # SECTION 1: CÁC TRANG SƠ KHỞI (MỤC LỤC, DANH MỤC BẢNG & HÌNH)
    # Số trang La Mã thường (i, ii, iii...) theo đúng QĐ 309
    # -----------------------------------------------------------------------
    sec_front = doc.add_section(WD_SECTION.NEW_PAGE)
    set_dnc_margins(sec_front)
    remove_pg_borders(sec_front)
    
    # Footer với số La Mã (i, ii, iii...)
    footer_front_p = sec_front.footer.paragraphs[0]
    footer_front_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_roman = footer_front_p.add_run()
    set_font(r_roman, 'Times New Roman', Pt(12))
    fld_roman = OxmlElement('w:fldSimple')
    fld_roman.set(qn('w:instr'), 'PAGE')
    r_roman._r.append(fld_roman)
    
    sectPr_front = sec_front._sectPr
    pgNumType_front = OxmlElement('w:pgNumType')
    pgNumType_front.set(qn('w:fmt'), 'lowerRoman')
    pgNumType_front.set(qn('w:start'), '1')
    sectPr_front.append(pgNumType_front)
    
    # Tiêu đề MỤC LỤC
    add_p(doc, "MỤC LỤC", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(12), space_after=Pt(10), first_line_indent=Cm(0))
    
    p_trang = doc.add_paragraph()
    p_trang.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_trang.paragraph_format.space_before = Pt(0)
    p_trang.paragraph_format.space_after = Pt(4)
    p_trang.paragraph_format.first_line_indent = Cm(0)
    r_trang = p_trang.add_run("Trang")
    set_font(r_trang, 'Times New Roman', Pt(13), italic=True)
    
    def add_toc_line(title, page_str, level=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(3) if level == 1 else Pt(1)
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
        elif level == 2:
            pf.left_indent = Cm(0.8)
            pf.first_line_indent = Cm(0)
            r1 = p.add_run(title)
            set_font(r1, 'Times New Roman', Pt(12.5), bold=False)
            r2 = p.add_run(f"\t{page_str}")
            set_font(r2, 'Times New Roman', Pt(12.5), bold=False)
        else:
            pf.left_indent = Cm(1.5)
            pf.first_line_indent = Cm(0)
            r1 = p.add_run(title)
            set_font(r1, 'Times New Roman', Pt(12), italic=True)
            r2 = p.add_run(f"\t{page_str}")
            set_font(r2, 'Times New Roman', Pt(12), italic=True)

    if tables:
        add_toc_line("DANH SÁCH BẢNG", "ii", level=1)
    if figures:
        add_toc_line("DANH SÁCH HÌNH", "iii" if tables else "ii", level=1)

    for item_title, item_page, item_lvl in toc_items:
        add_toc_line(item_title, str(item_page), level=item_lvl)

    # Danh sách Bảng nếu có trong nội dung
    if tables:
        doc.add_page_break()
        add_p(doc, "DANH SÁCH BẢNG", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=Pt(12), space_after=Pt(10), first_line_indent=Cm(0))
        for tbl_title, tbl_page in tables:
            add_toc_line(tbl_title, str(tbl_page), level=1)

    # Danh sách Hình nếu có trong nội dung
    if figures:
        doc.add_page_break()
        add_p(doc, "DANH SÁCH HÌNH", font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
              space_before=Pt(12), space_after=Pt(10), first_line_indent=Cm(0))
        for fig_title, fig_page in figures:
            add_toc_line(fig_title, str(fig_page), level=1)

    # -----------------------------------------------------------------------
    # SECTION 2: THÂN BÀI BÁO CÁO (Chương 1 trở đi)
    # Số trang Ả Rập (1, 2, 3...) bắt đầu từ 1 theo đúng QĐ 309
    # -----------------------------------------------------------------------
    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    set_dnc_margins(sec_body)
    remove_pg_borders(sec_body)
    
    # Thiết lập số trang Ả Rập bắt đầu từ 1 ở đầu thân bài
    sectPr_body = sec_body._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:fmt'), 'decimal')
    pgNumType.set(qn('w:start'), '1')
    sectPr_body.append(pgNumType)
    
    # Header & Footer
    footer = sec_body.footer
    p_ft = footer.paragraphs[0]
    p_ft.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ft = p_ft.add_run()
    set_font(r_ft, 'Times New Roman', Pt(12))
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), 'PAGE')
    r_ft._r.append(fld)
    
    lines = raw_text.split('\n')
    is_first_chapter_rendered = False
    for line in lines:
        line_s = line.strip()
        if not line_s:
            continue
            
        # 1. Phát hiện Tên Chương
        if re.match(r'^(CHƯƠNG|Chương)\s+\d+', line_s, re.IGNORECASE):
            if is_first_chapter_rendered:
                doc.add_page_break()
            else:
                is_first_chapter_rendered = True
            add_p(doc, line_s.upper(), font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_before=Pt(18), space_after=Pt(12), first_line_indent=Cm(0))
            continue
            
        # 2. Phát hiện Tiêu đề lớn (TÀI LIỆU THAM KHẢO, PHỤ LỤC...)
        if re.match(r'^(TÀI LIỆU THAM KHẢO|PHỤ LỤC|KẾT LUẬN VÀ KIẾN NGHỊ)', line_s, re.IGNORECASE):
            doc.add_page_break()
            add_p(doc, line_s.upper(), font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_before=Pt(18), space_after=Pt(12), first_line_indent=Cm(0))
            continue

        # 3. Phát hiện Mục cấp 1 (ví dụ: 1.1, 2.1)
        if re.match(r'^\d+\.\d+\s+[^\.]', line_s):
            add_p(doc, line_s.upper(), font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=Pt(12), space_after=Pt(4), first_line_indent=Cm(0))
            continue
            
        # 4. Phát hiện Mục cấp 2 (ví dụ: 1.1.1, 2.1.1)
        if re.match(r'^\d+\.\d+\.\d+\s+[^\.]', line_s):
            add_p(doc, line_s, font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=Pt(8), space_after=Pt(3), first_line_indent=Cm(1.0))
            continue

        # 5. Phát hiện Tiêu đề Hình / Bảng
        if re.match(r'^(Hình|Bảng|Fig\.|Table)\s+\d+', line_s, re.IGNORECASE):
            add_p(doc, line_s, font_size=Pt(13), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                  space_before=Pt(4), space_after=Pt(6), first_line_indent=Cm(0))
            continue

        # 6. Phát hiện danh sách đầu dòng (- hoặc *)
        if re.match(r'^[\-\*\•]\s+', line_s):
            add_p(doc, line_s, font_size=Pt(13), bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              space_before=Pt(2), space_after=Pt(2), line_spacing=1.2,
              left_indent=Cm(1.5), first_line_indent=Cm(-0.5))
            continue

        # 7. Đoạn văn nội dung thông thường
        add_p(doc, line_s, font_size=Pt(13), bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              space_before=Pt(6), space_after=Pt(0), line_spacing=1.2,
              first_line_indent=Cm(1.0))

    doc.save(output_docx)
    return output_docx

def parse_and_format_springer(raw_text, meta, output_docx):
    doc = docx.Document()
    
    # Section 0: Trang bìa DNC (đáp ứng 4 nội dung môn học, lớp, GV, kỳ)
    add_dnc_cover(doc, meta, is_midterm=True)
    
    # Section 1: Bài báo Springer LNCS
    sec_sp = doc.add_section(WD_SECTION.NEW_PAGE)
    set_springer_margins(sec_sp)
    remove_pg_borders(sec_sp)
    sec_sp.different_first_page_header_footer = True
    
    # Enable even/odd headers
    settings = doc.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        evenAndOdd = OxmlElement('w:evenAndOddHeaders')
        settings.append(evenAndOdd)
        
    sec_sp.header.is_linked_to_previous = False
    sec_sp.even_page_header.is_linked_to_previous = False
    sec_sp.footer.is_linked_to_previous = False
    
    text_width_cm = 12.2
    
    # Odd Header
    odd_p = sec_sp.header.paragraphs[0]
    odd_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    odd_p.paragraph_format.tab_stops.add_tab_stop(Cm(text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
    short_title = meta.get("ten_de_tai_tieng_anh", "Research Paper")[:60]
    r_odd_title = odd_p.add_run(f"{short_title}\t")
    set_font(r_odd_title, 'Times New Roman', Pt(8.5))
    r_odd_page = odd_p.add_run()
    set_font(r_odd_page, 'Times New Roman', Pt(9))
    fld_odd = OxmlElement('w:fldSimple')
    fld_odd.set(qn('w:instr'), 'PAGE')
    r_odd_page._r.append(fld_odd)
    
    # Even Header
    even_p = sec_sp.even_page_header.paragraphs[0]
    even_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    even_p.paragraph_format.tab_stops.add_tab_stop(Cm(text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
    r_even_page = even_p.add_run()
    set_font(r_even_page, 'Times New Roman', Pt(9))
    fld_even = OxmlElement('w:fldSimple')
    fld_even.set(qn('w:instr'), 'PAGE')
    r_even_page._r.append(fld_even)
    author_short = f"\t{meta.get('ho_ten', 'Author')} and {meta.get('giang_vien', 'Advisor')}"
    r_even_author = even_p.add_run(author_short)
    set_font(r_even_author, 'Times New Roman', Pt(8.5))

    sectPr_sp = sec_sp._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:fmt'), 'decimal')
    pgNumType.set(qn('w:start'), '1')
    sectPr_sp.append(pgNumType)

    # Title bài báo
    en_title = meta.get("ten_de_tai_tieng_anh", meta.get("ten_de_tai", "Research Title"))
    add_p(doc, en_title, font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(18), space_after=Pt(12), line_spacing=1.15, first_line_indent=Cm(0))

    # Tác giả & Giảng viên
    author_line = f"{meta.get('ho_ten', 'Author')}¹ and {meta.get('giang_vien', 'Advisor')}¹*"
    add_p(doc, author_line, font_size=Pt(10), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(4), first_line_indent=Cm(0))
    
    # Đơn vị
    add_p(doc, f"¹ Faculty of Information Technology, Nam Can Tho University, Vietnam",
          font_size=Pt(9), italic=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(14), first_line_indent=Cm(0))

    lines = raw_text.split('\n')
    for line in lines:
        line_s = line.strip()
        if not line_s:
            continue
            
        # Section level 1 (e.g., "1 Introduction")
        if re.match(r'^\d+\s+[A-Z]', line_s):
            add_p(doc, line_s, font_size=Pt(12), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=Pt(14), space_after=Pt(6), first_line_indent=Cm(0))
            continue
            
        # Section level 2 (e.g., "2.1 Background")
        if re.match(r'^\d+\.\d+\s+[A-Z]', line_s):
            add_p(doc, line_s, font_size=Pt(10), bold=True, align=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=Pt(10), space_after=Pt(4), first_line_indent=Cm(0))
            continue

        # Abstract & Keywords
        if line_s.startswith("Abstract") or line_s.startswith("Keywords"):
            p = add_p(doc, line_s, font_size=Pt(9), bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                      space_before=Pt(0), space_after=Pt(8), line_spacing=1.0,
                      left_indent=Cm(1.0), right_indent=Cm(1.0), first_line_indent=Cm(0))
            continue

        # Normal paragraph (Springer style: 10pt, justify, single line, indent 0.5cm)
        add_p(doc, line_s, font_size=Pt(10), bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              space_before=Pt(0), space_after=Pt(0), line_spacing=1.0,
              first_line_indent=Cm(0.5))

    doc.save(output_docx)
    return output_docx

def convert_docx_to_pdf_win32(docx_path, pdf_path):
    import win32com.client
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(abs_docx)
        doc.ExportAsFixedFormat(abs_pdf, 17) # wdExportFormatPDF = 17
        doc.Close(0)
        return True
    finally:
        word.Quit()

