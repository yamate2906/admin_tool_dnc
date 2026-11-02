import win32com.client
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

def convert_all():
    files = [
        ("Bao_cao_Giua_ki_Springer_Dinh_Van_Lam.docx", "Bao_cao_Giua_ki_Springer_Dinh_Van_Lam.pdf"),
        ("Bao_cao_Cuoi_ki_DNC_309_Dinh_Van_Lam.docx", "Bao_cao_Cuoi_ki_DNC_309_Dinh_Van_Lam.pdf")
    ]
    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for docx_file, pdf_file in files:
            abs_docx = os.path.abspath(docx_file)
            abs_pdf = os.path.abspath(pdf_file)
            print(f"Opening: {abs_docx}")
            doc = word.Documents.Open(abs_docx)
            print(f"Saving PDF: {abs_pdf}")
            # 17 = wdFormatPDF
            doc.SaveAs(abs_pdf, FileFormat=17)
            doc.Close(False)
            print(f"Created {pdf_file}: {os.path.getsize(abs_pdf):,} bytes")
    finally:
        word.Quit()
    print("All conversions completed successfully.")

if __name__ == "__main__":
    convert_all()

