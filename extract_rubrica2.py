import PyPDF2
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

r = PyPDF2.PdfReader(r'C:/Users/ignac/Downloads/Rubrica de Logro_1ACC0263_2026-10-(TF).pdf')
for i, page in enumerate(r.pages):
    text = page.extract_text()
    if text:
        print(f"=== PAGE {i+1} ===")
        print(text[:3000])
        print()
