import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
import shutil
import os

def create_springer_document():
    # Read existing paragraphs from Random_Forest_WebGIS_Expanded.docx
    source_doc = docx.Document('Random_Forest_WebGIS_Expanded.docx')
    raw_paragraphs = [p.text.strip() for p in source_doc.paragraphs if p.text.strip()]
    
    doc = docx.Document()
    
    # -------------------------------------------------------------
    # SECTION 0: TRANG BÌA (CHUẨN DNC 309 + KHOA CNTT + GIẢNG VIÊN)
    # -------------------------------------------------------------
    sec_cover = doc.sections[0]
    sec_cover.page_width = Cm(21.0)
    sec_cover.page_height = Cm(29.7)
    sec_cover.left_margin = Cm(3.5)
    sec_cover.right_margin = Cm(2.0)
    sec_cover.top_margin = Cm(2.5)
    sec_cover.bottom_margin = Cm(2.5)
    sec_cover.header_distance = Cm(1.0)
    sec_cover.footer_distance = Cm(1.0)
    
    # Decorative border for outer cover page only
    sectPr_cover = sec_cover._sectPr
    pgBorders = OxmlElement('w:pgBorders')
    pgBorders.set(qn('w:display'), 'firstPage')
    for side in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'twistedLines1')
        b.set(qn('w:sz'), '18')
        b.set(qn('w:space'), '4' if side == 'left' else '2')
        b.set(qn('w:color'), 'auto')
        pgBorders.append(b)
    sectPr_cover.append(pgBorders)
    
    def add_cp(text='', font_size=Pt(13), bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.CENTER,
               space_before=Pt(0), space_after=Pt(0), line_spacing=1.15):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = space_before
        pf.space_after = space_after
        pf.line_spacing = line_spacing
        pf.first_line_indent = Cm(0)
        if text:
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = font_size
            r.bold = bold
            r.italic = italic
        return p

    # 1. Cơ quan chủ quản và Khoa (14pt, QĐ 309 & Mẫu khoa)
    add_cp("TRƯỜNG ĐẠI HỌC NAM CẦN THƠ", font_size=Pt(14), bold=False, space_before=Pt(6), space_after=Pt(3))
    add_cp("KHOA CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, space_after=Pt(10))

    # 2. Logo chính thức DNC (theo mẫu 3 Mau bao cao DO AN_2023.docx và hình ảnh Word)
    if os.path.exists("dnc_logo.png"):
        p_logo = doc.add_paragraph()
        p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_logo.paragraph_format.space_before = Pt(4)
        p_logo.paragraph_format.space_after = Pt(14)
        p_logo.paragraph_format.first_line_indent = Cm(0)
        r_logo = p_logo.add_run()
        r_logo.add_picture("dnc_logo.png", width=Cm(4.2))

    # 3. Khối Tác giả sinh viên (14pt in đậm)
    add_cp("ĐINH VĂN LẮM", font_size=Pt(14), bold=True, space_after=Pt(3))
    add_cp("MSSV: 235292", font_size=Pt(14), bold=True, space_after=Pt(3))
    add_cp("LỚP: DH23TIN06", font_size=Pt(14), bold=True, space_after=Pt(16))

    # 4. Khối Đề tài (Cỡ chữ 20 in đậm theo đúng QĐ 309 Mục 2.1)
    add_cp("ĐỀ TÀI", font_size=Pt(14), bold=True, space_after=Pt(6))
    add_cp("XÂY DỰNG HỆ THỐNG WEBGIS DỰ ĐOÁN GIAO THÔNG\nBẰNG RANDOM FOREST",
           font_size=Pt(20), bold=True, space_after=Pt(5), line_spacing=1.2)
    add_cp("(AN OVERVIEW STUDY ON THE APPLICATION OF THE RANDOM FOREST ALGORITHM\nAND WEBGIS PLATFORM IN TRAFFIC FORECASTING)",
           font_size=Pt(11.5), italic=True, space_after=Pt(18), line_spacing=1.15)

    # 5. Khối Báo cáo giữa kỳ & Giảng viên hướng dẫn (Thể hiện đủ 4 nội dung GV yêu cầu)
    add_cp("BÁO CÁO GIỮA KỲ", font_size=Pt(14), bold=True, space_after=Pt(3))
    add_cp("Phương Pháp Nghiên Cứu và Viết Báo Cáo Khoa Học", font_size=Pt(13), bold=True, space_after=Pt(3))
    add_cp("Học kỳ I", font_size=Pt(13), bold=True, space_after=Pt(3))
    add_cp("GIẢNG VIÊN HƯỚNG DẪN: TS. NGÔ HỒ ANH KHÔI", font_size=Pt(13), bold=True, space_after=Pt(3))
    add_cp("NGÀNH: CÔNG NGHỆ THÔNG TIN", font_size=Pt(14), bold=True, space_after=Pt(3))
    add_cp("Mã số ngành: 7480201", font_size=Pt(14), bold=True, space_after=Pt(26))

    # 6. Địa điểm và thời gian (13pt regular)
    add_cp("Cần Thơ, Tháng 11/2026", font_size=Pt(13), bold=False, space_after=Pt(0))

    # -------------------------------------------------------------
    # SECTION 1: BÀI BÁO SPRINGER LNCS (Bắt đầu từ Section mới)
    # HOÀN TOÀN KHÔNG CÓ KHUNG VIỀN TRANG (REMOVE PGBORDERS)
    # -------------------------------------------------------------
    sec_springer = doc.add_section(WD_SECTION.NEW_PAGE)
    
    # Remove pgBorders so Springer paper has NO frame!
    sectPr_springer = sec_springer._sectPr
    b_sp = sectPr_springer.find(qn('w:pgBorders'))
    if b_sp is not None:
        sectPr_springer.remove(b_sp)
        
    sec_springer.different_first_page_header_footer = True
    
    # Page setup - Springer LNCS standard:
    sec_springer.page_width = Cm(21.0)
    sec_springer.page_height = Cm(29.7)
    sec_springer.left_margin = Cm(4.4)
    sec_springer.right_margin = Cm(4.4)
    sec_springer.top_margin = Cm(4.5)
    sec_springer.bottom_margin = Cm(4.5)
    sec_springer.header_distance = Cm(3.2)
    sec_springer.footer_distance = Cm(3.2)
    
    text_width_cm = 12.2
    
    # Enable even and odd headers in document settings
    settings = doc.settings.element
    if settings.find(qn('w:evenAndOddHeaders')) is None:
        evenAndOdd = OxmlElement('w:evenAndOddHeaders')
        settings.append(evenAndOdd)
        
    sec_springer.header.is_linked_to_previous = False
    sec_springer.even_page_header.is_linked_to_previous = False
    sec_springer.footer.is_linked_to_previous = False
    
    # Odd Page Header (Short Title on Left, Page number on Right)
    odd_p = sec_springer.header.paragraphs[0]
    odd_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    odd_p.paragraph_format.tab_stops.add_tab_stop(Cm(text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
    r_odd_title = odd_p.add_run("Overview Study on Random Forest and WebGIS in Traffic Forecasting\t")
    r_odd_title.font.name = 'Times New Roman'
    r_odd_title.font.size = Pt(8.5)
    r_odd_page = odd_p.add_run()
    r_odd_page.font.name = 'Times New Roman'
    r_odd_page.font.size = Pt(9)
    fld_odd = OxmlElement('w:fldSimple')
    fld_odd.set(qn('w:instr'), 'PAGE')
    r_odd_page._r.append(fld_odd)
    
    # Even Page Header (Page number on Left, Author on Right)
    even_p = sec_springer.even_page_header.paragraphs[0]
    even_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    even_p.paragraph_format.tab_stops.add_tab_stop(Cm(text_width_cm), WD_TAB_ALIGNMENT.RIGHT)
    r_even_page = even_p.add_run()
    r_even_page.font.name = 'Times New Roman'
    r_even_page.font.size = Pt(9)
    fld_even = OxmlElement('w:fldSimple')
    fld_even.set(qn('w:instr'), 'PAGE')
    r_even_page._r.append(fld_even)
    r_even_author = even_p.add_run("\tD. V. Lam and A. K. Ngo-Ho")
    r_even_author.font.name = 'Times New Roman'
    r_even_author.font.size = Pt(8.5)

    # Set page number starting at 1 for the Springer paper
    pgNumType_springer = OxmlElement('w:pgNumType')
    pgNumType_springer.set(qn('w:fmt'), 'decimal')
    pgNumType_springer.set(qn('w:start'), '1')
    sectPr_springer.append(pgNumType_springer)

    def add_p(text, font_name='Times New Roman', font_size=Pt(10),
              bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=Pt(0), space_after=Pt(0), line_spacing=1.0,
              first_line_indent=Cm(0), left_indent=Cm(0), right_indent=Cm(0)):
        p = doc.add_paragraph()
        p.alignment = align
        p_format = p.paragraph_format
        p_format.space_before = space_before
        p_format.space_after = space_after
        p_format.line_spacing = line_spacing
        p_format.first_line_indent = first_line_indent
        p_format.left_indent = left_indent
        p_format.right_indent = right_indent
        
        run = p.add_run(text)
        run.font.name = font_name
        run.font.size = font_size
        run.bold = bold
        run.italic = italic
        return p

    def add_image_figure(image_path, caption_text, width_cm=11.5):
        if os.path.exists(image_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(3)
            p_img.paragraph_format.first_line_indent = Cm(0)
            r_img = p_img.add_run()
            r_img.add_picture(image_path, width=Cm(width_cm))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            p_cap.paragraph_format.line_spacing = 1.0
            p_cap.paragraph_format.first_line_indent = Cm(0)
            
            r_c1 = p_cap.add_run(caption_text.split('.')[0] + '.')
            r_c1.font.name = 'Times New Roman'
            r_c1.font.size = Pt(9)
            r_c1.bold = True
            
            rest = '.'.join(caption_text.split('.')[1:])
            if rest:
                r_c2 = p_cap.add_run(rest)
                r_c2.font.name = 'Times New Roman'
                r_c2.font.size = Pt(9)

    # Title
    add_p("An Overview Study on the Application of the Random Forest Algorithm and WebGIS Platform in Traffic Forecasting",
          font_size=Pt(14), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(18), space_after=Pt(14), line_spacing=1.15)
    
    # Author with Advisor (TS. Ngô Hồ Anh Khôi)
    add_p("Dinh-Van LAM¹[0009-0002-3529-2000] and Anh-Khoi NGO-HO¹*[0000-0001-5444-5927]",
          font_size=Pt(10), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(4))
    
    # Affiliation
    add_p("¹ Faculty of Information Technology, Nam Can Tho University,\n168 Nguyen Van Cu Street, An Binh Ward, Can Tho City, Vietnam",
          font_size=Pt(9), italic=False, align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(4), line_spacing=1.1)
    
    # Email
    add_p("{s235292@nctu.edu.vn, nhakhoi@nctu.edu.vn, dinhlamtq2000@gmail.com}",
          font_name='Courier New', font_size=Pt(9), align=WD_ALIGN_PARAGRAPH.CENTER,
          space_before=Pt(0), space_after=Pt(16))
    
    # Abstract
    abstract_raw = raw_paragraphs[4]
    abstract_text = abstract_raw[len("Abstract"):].strip() if abstract_raw.startswith("Abstract") else abstract_raw
        
    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs.paragraph_format.space_before = Pt(0)
    p_abs.paragraph_format.space_after = Pt(6)
    p_abs.paragraph_format.line_spacing = 1.0
    p_abs.paragraph_format.left_indent = Cm(1.0)
    p_abs.paragraph_format.right_indent = Cm(1.0)
    
    r_abs_lbl = p_abs.add_run("Abstract. ")
    r_abs_lbl.font.name = 'Times New Roman'
    r_abs_lbl.font.size = Pt(9)
    r_abs_lbl.bold = True
    
    r_abs_body = p_abs.add_run(abstract_text)
    r_abs_body.font.name = 'Times New Roman'
    r_abs_body.font.size = Pt(9)
    
    # Keywords
    kw_raw = raw_paragraphs[5]
    kw_text = kw_raw[len("Keywords:"):].strip() if kw_raw.startswith("Keywords:") else kw_raw
    kw_formatted = kw_text.replace(",", " \u00b7")
    
    p_kw = doc.add_paragraph()
    p_kw.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_kw.paragraph_format.space_before = Pt(0)
    p_kw.paragraph_format.space_after = Pt(18)
    p_kw.paragraph_format.line_spacing = 1.0
    p_kw.paragraph_format.left_indent = Cm(1.0)
    p_kw.paragraph_format.right_indent = Cm(1.0)
    
    r_kw_lbl = p_kw.add_run("Keywords: ")
    r_kw_lbl.font.name = 'Times New Roman'
    r_kw_lbl.font.size = Pt(9)
    r_kw_lbl.bold = True
    
    r_kw_body = p_kw.add_run(kw_formatted)
    r_kw_body.font.name = 'Times New Roman'
    r_kw_body.font.size = Pt(9)

    def add_section_h1(num, title):
        add_p(f"{num}  {title}", font_size=Pt(12), bold=True,
              space_before=Pt(14), space_after=Pt(6))
        
    def add_section_h2(num_str, title):
        add_p(f"{num_str}  {title}", font_size=Pt(10), bold=True,
              space_before=Pt(10), space_after=Pt(4))
        
    def add_body_paragraph(text, is_first_after_heading=False):
        first_indent = Cm(0) if is_first_after_heading else Cm(0.5)
        add_p(text, font_size=Pt(10), align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              space_before=Pt(0), space_after=Pt(0), line_spacing=1.0,
              first_line_indent=first_indent)

    # 1 Introduction
    add_section_h1(1, "Introduction")
    add_body_paragraph(raw_paragraphs[7], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[8])
    add_body_paragraph(raw_paragraphs[9])
    add_body_paragraph(raw_paragraphs[10])
    add_body_paragraph(raw_paragraphs[11])
    
    # 2 Overview of GIS Product Platforms
    add_section_h1(2, "Overview of Geographic Information System Product Platforms")
    add_section_h2("2.1", "Commercial Mapping Platforms and Cloud Solutions")
    add_body_paragraph(raw_paragraphs[13], is_first_after_heading=True)
    add_section_h2("2.2", "Open-Source Spatial Visualization Ecosystems")
    add_body_paragraph(raw_paragraphs[14], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[15])
    add_body_paragraph(raw_paragraphs[16])
    add_section_h2("2.3", "Web Architecture and Spatial Data Interoperability")
    add_body_paragraph(raw_paragraphs[17], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[18])
    add_body_paragraph(raw_paragraphs[19])
    
    # Insert Figure 1 & Figure 2
    add_image_figure('figures/fig_2_1_architecture.png', 'Fig. 1. Multi-tier WebGIS architectural framework for intelligent traffic volume forecasting.')
    add_image_figure('figures/fig_2_2_ogc_flow.png', 'Fig. 2. Spatial data interoperability pipeline based on Open Geospatial Consortium (OGC) specifications.')
    
    # 3 Source Codes and Machine Learning Libraries
    add_section_h1(3, "Source Codes and Machine Learning Libraries")
    add_section_h2("3.1", "Algorithmic Foundation and Computational Stack")
    add_body_paragraph(raw_paragraphs[21], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[22])
    add_section_h2("3.2", "Random Forest Regressor Architecture and Hyperparameters")
    add_body_paragraph(raw_paragraphs[23], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[24])
    
    # Insert Figure 3 & Figure 4
    add_image_figure('figures/fig_3_1_random_forest.png', 'Fig. 3. Ensemble aggregation and bootstrap sampling mechanics in Random Forest regression.')
    
    add_section_h2("3.3", "Pipeline Design, Validation, and Operational Lifecycles")
    add_body_paragraph(raw_paragraphs[25], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[26])
    add_body_paragraph(raw_paragraphs[27])
    add_body_paragraph(raw_paragraphs[28])
    add_image_figure('figures/fig_3_2_pipeline.png', 'Fig. 4. End-to-end data preprocessing and model estimation pipeline.')
    
    # 4 Training Datasets
    add_section_h1(4, "Training Datasets")
    add_section_h2("4.1", "Benchmark Profile and Data Granularity")
    add_body_paragraph(raw_paragraphs[30], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[31])
    add_body_paragraph(raw_paragraphs[32])
    add_section_h2("4.2", "Exploratory Profiling and Feature Preprocessing")
    add_body_paragraph(raw_paragraphs[33], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[34])
    add_section_h2("4.3", "Spatial-Temporal Coverage and Chronological Partitioning")
    add_body_paragraph(raw_paragraphs[35], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[36])
    add_body_paragraph(raw_paragraphs[37])
    
    # Insert Figure 5
    add_image_figure('figures/fig_4_1_traffic_pattern.png', 'Fig. 5. Average hourly traffic volume distribution on weekdays and weekends.')
    
    # 5 Measurement and Model Evaluation
    add_section_h1(5, "Measurement and Model Evaluation")
    add_section_h2("5.1", "Mathematical Formulations of Evaluation Metrics")
    add_body_paragraph(raw_paragraphs[39], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[40])
    add_body_paragraph(raw_paragraphs[41])
    add_body_paragraph(raw_paragraphs[42])
    add_body_paragraph(raw_paragraphs[43])
    add_section_h2("5.2", "Real-World Workflow Simulation and Operational Bounds")
    add_body_paragraph(raw_paragraphs[44], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[45])
    add_body_paragraph(raw_paragraphs[46])
    
    # Insert Figure 6
    add_image_figure('figures/fig_4_2_prediction_comparison.png', 'Fig. 6. Continuous 48-hour time series trajectory: ground truth vs. Random Forest predictions.')
    
    # 6 History of the Topic and Major Research
    add_section_h1(6, "History of the Topic and Major Research")
    add_section_h2("6.1", "Evolutionary Trajectory of Ensemble Traffic Forecasting")
    add_body_paragraph(raw_paragraphs[48], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[49])
    add_body_paragraph(raw_paragraphs[50])
    add_section_h2("6.2", "Spatiotemporal Paradigms and WebGIS Evolution")
    add_body_paragraph(raw_paragraphs[51], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[52])
    add_body_paragraph(raw_paragraphs[53])
    add_section_h2("6.3", "Identified Literature Gaps and Research Synthesis")
    add_body_paragraph(raw_paragraphs[54], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[55])
    
    # 7 Conclusion
    add_section_h1(7, "Conclusion")
    add_section_h2("7.1", "Synthesis of Critical Insights")
    add_body_paragraph(raw_paragraphs[57], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[58])
    add_body_paragraph(raw_paragraphs[59])
    add_body_paragraph(raw_paragraphs[60])
    add_section_h2("7.2", "Practical Implementation Roadmap and Outlook")
    add_body_paragraph(raw_paragraphs[61], is_first_after_heading=True)
    add_body_paragraph(raw_paragraphs[62])
    add_body_paragraph(raw_paragraphs[63])
    add_body_paragraph(raw_paragraphs[64])
    
    # References
    add_p("References", font_size=Pt(12), bold=True,
          space_before=Pt(16), space_after=Pt(8))
    
    ref_items = raw_paragraphs[66:83]
    for idx, ref in enumerate(ref_items, start=1):
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.space_before = Pt(0)
        p_ref.paragraph_format.space_after = Pt(3)
        p_ref.paragraph_format.line_spacing = 1.0
        p_ref.paragraph_format.left_indent = Cm(0.5)
        p_ref.paragraph_format.first_line_indent = Cm(-0.5)
        
        r_num = p_ref.add_run(f"{idx}. ")
        r_num.font.name = 'Times New Roman'
        r_num.font.size = Pt(9)
        r_num.bold = False
        
        r_text = p_ref.add_run(ref)
        r_text.font.name = 'Times New Roman'
        r_text.font.size = Pt(9)
        
    out_path_docx = "Bao_cao_Giua_ki_Springer_Dinh_Van_Lam.docx"
    doc.save(out_path_docx)
    print(f"Generated {out_path_docx} successfully.")
    
    out_path_doc = "Bao_cao_Giua_ki_Springer_Dinh_Van_Lam.doc"
    shutil.copyfile(out_path_docx, out_path_doc)
    print(f"Copied to {out_path_doc}.")

if __name__ == "__main__":
    create_springer_document()
