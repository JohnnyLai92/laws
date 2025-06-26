import os
from PyPDF2 import PdfReader
from openpyxl import Workbook

folder = '雲林縣政府法規'

wb = Workbook()
ws = wb.active
ws.append(['檔名', '文件頁數', '字數', '含圖表', '需OCR', '常見問答1', '常見問答2', '常見問答3'])

for filename in os.listdir(folder):
    if not filename.lower().endswith('.pdf'):
        continue
    path = os.path.join(folder, filename)
    reader = PdfReader(path)
    text = ''
    for page in reader.pages:
        if page.extract_text():
            text += page.extract_text()
    num_pages = len(reader.pages)
    num_words = len(text.split())
    has_table = 'Table' in text or '表' in text
    needs_ocr = not bool(text.strip())
    qas = [f"{filename} 的目的為何?", f"{filename} 的適用對象是誰?", f"{filename} 的施行日期是什麼時候?"]
    ws.append([filename, num_pages, num_words, '是' if has_table else '否', '是' if needs_ocr else '否'] + qas)

wb.save('lawsinfo.xlsx')
print('資料已寫入 lawsinfo.xlsx')
