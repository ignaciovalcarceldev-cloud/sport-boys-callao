import PyPDF2
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

r = PyPDF2.PdfReader(r'C:/Users/ignac/Downloads/1AAD2788 Avance de Trabajo (TB1) - modificado.pdf')
for i, page in enumerate(r.pages[3:]):
    text = page.extract_text()
    if text:
        print(f"=== PAGE {i+4} ===")
        print(text[:3000])
        print()
