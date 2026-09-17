import zipfile
import xml.etree.ElementTree as ET

z = zipfile.ZipFile(r'C:/Users/ignac/Downloads/Investigacion_Startup_IA_Peru.docx')
tree = ET.parse(z.open('word/document.xml'))
root = tree.getroot()
ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
text = '\n'.join([p.text or '' for p in root.iter(ns + 't')])
print(text)
