import re
import pdfplumber
import pandas as pd
from datetime import datetime
import fitz
import os
from PIL import Image
import pytesseract
import xlsxwriter
from functions import clean_amount, copy_format

balance_sheet = 'temp/balance_sheet2.pdf'
income_statement_sheet = 'temp/income_statement2.pdf'
assets = []
liabilities = []
equities = []
retained_earnings = []

# set ocr path in windows 
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
is_revenue = []
is_cos = []
is_oe = []

# ============================================ Balance Sheet ==================================
# no layout to get the first line, company name, date
text_no_layout = ''
with pdfplumber.open(balance_sheet) as pdf:
    page = pdf.pages
    for page in pdf.pages:
        single_page_text = page.extract_text()
        text_no_layout = text_no_layout + '\n' + single_page_text

# get the company name, year end, current year and previous year
first_line = text_no_layout.split('\n', 2)[1]
company_name = first_line.split('   ', 1)[0]
date = first_line.split('   ', 2)[2]
match = re.search(r'\d{4}-\d{2}-\d{2}', date)
extracted_date = datetime.strptime(match.group(), '%Y-%m-%d').date()
final_date = extracted_date.strftime("%b %d, %Y")
year = extracted_date.strftime("%Y")

# text with layout to get the body of the document
text = '' 
with pdfplumber.open(balance_sheet) as pdf:
    page = pdf.pages
    for page in pdf.pages:
        single_page_text = page.extract_text(layout=True)
        text = text + '\n' + single_page_text

# get the assets from balance sheet
assets_re = re.compile(r'(?:\bAssets\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bLiabilities\s.*Code\s.* Current\s.*year\s.*Prior\s.*year\b)')
assets_line = assets_re.findall(text)
if assets_line:
    for line in assets_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,8}(\(?\d{1,3},?\d{3}\)?)?(\s{3,12})?(\(?\d{1,3},?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                assets.append(item)

# get the liabilities from balance sheet
liabilities_re = re.compile(r'(?:\bLiabilities\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bEquity\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)')
liabilities_line = liabilities_re.findall(text)
if liabilities_line:
    for line in liabilities_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,8}(\(?\d{1,3},?\d{3}\)?)?(\s{3,12})?(\(?\d{1,3},?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                liabilities.append(item)

# get the equities from balance sheet
equity_re = re.compile(r'(?:\bEquity\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bRetained\s.*earnings\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)')
equity_line = equity_re.findall(text)
if equity_line:
    for line in equity_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,8}(\(?\d{1,3},?\d{3}\)?)?(\s{3,12})?(\(?\d{1,3},?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                equities.append(item)

# get the retained_earnings from balance sheet
retained_re = re.compile(r'(?:\bRetained\s.*earnings\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\b\\*The\s.*amount\s.*on\s.*line\b)')
retained_line = retained_re.findall(text)
if retained_line:
    for line in retained_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,8}(\(?\d{1,3},?\d{3}\)?)?(\s{3,12})?(\(?\d{1,3},?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                retained_earnings.append(item)

# ============================================ Income Statement Sheet ==================================
zoom = 3

# convert income statement sheet to images
with fitz.open(income_statement_sheet) as doc:
    digits = len(str(doc.page_count))
    for i, page in enumerate(doc):
        if doc.page_count>1:
            page_id = f'{{:0{digits}}}'.format(i+1)
            file_name = f'{os.path.split(income_statement_sheet)[1][:-4]}_p{page_id}.png'
        else:
            file_name = f'{os.path.split(income_statement_sheet)[1][:-4]}.png'
        png_file = os.path.join('temp/', file_name)
        trans = fitz.Matrix(zoom, zoom).prerotate(0) # zoom_x, zoom_y
        pm = page.get_pixmap(matrix=trans, alpha=False)
        pm.save(png_file)

        # convert image to searchable pdf
        is_pdf_p = pytesseract.image_to_pdf_or_hocr('temp/' + file_name, extension='pdf')
        with open(f'temp/is_pdf_p{page_id}.pdf', 'w+b') as f:
            f.write(is_pdf_p) 
    
# get text from pdf pages
is_text_p1 = ''
is_text_p2 = ''
with pdfplumber.open("temp/is_pdf_p1.pdf") as pdf:
    page = pdf.pages
    for page in pdf.pages:
        single_page_text = page.extract_text(layout=True)
        is_text_p1 = is_text_p1 + '\n' + single_page_text
with pdfplumber.open("temp/is_pdf_p2.pdf") as pdf:
    page = pdf.pages
    for page in pdf.pages:
        single_page_text = page.extract_text(layout=False)
        is_text_p2 = is_text_p2 + '\n' + single_page_text

is_revenue_re = re.compile(r'(?:\bRevenue\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bCost\s.*of\s.*sales\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)')
is_revenue_line = is_revenue_re.findall(is_text_p1)
if is_revenue_line:
    for line in is_revenue_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,30}(\(?(\d{1,3})?,?\d{3}\)?)?(\s{3,30})?(\(?(\d{1,3})?,?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                is_revenue.append(item)

# get is Cost of sales
is_cos_re = re.compile(r'(?:\bCost\s.*of\s.*sales\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bOperating\s.*expenses\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)')
is_cos_line = is_cos_re.findall(is_text_p1)
if is_cos_line:
    for line in is_cos_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,30}(\(?(\d{1,3})?,?\d{3}\)?)?(\s{3,30})?(\(?(\d{1,3})?,?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                is_cos.append(item)

# get is Operating expenses
is_oe_re = re.compile(r'(?:\bOperating\s.*expenses\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)\n([\s\S]*)(?:\bFarming\s.*revenue\s.*Code\s.*Current\s.*year\s.*Prior\s.*year\b)')
is_oe_line = is_oe_re.findall(is_text_p1)
if is_oe_line:
    for line in is_oe_line[0].split('\n'):
        single_line_re = re.compile(r'(\w.+)(\d{4})\s{6,30}(\(?(\d{1,3})?,?\d{3}\)?)?(\s{3,30})?(\(?(\d{1,3})?,?\d{3}\)?)?')
        single_line = single_line_re.findall(line)
        if single_line:
            for item in single_line:
                is_oe.append(item)

# get is taxes and net income
is_current_taxes_re = re.compile(r'(\bCurrent\s.*income\s.*taxes\b)\s{2,30}(\d{4})\s{2,30}\-?\+?\=?\s{2,30}(\(?(\d{1,3})?,?\d{3}\)?)?\s{2,30}\-?\+?\=?\s{2,30}(\(?(\d{1,3})?,?\d{3}\)?)?')
is_current_taxes = is_current_taxes_re.findall(is_text_p2)
is_net_income_re = re.compile(r'(\bNet\s.*income\s.*loss\s.*after\s.*taxes\s.*and\s.*extraordinary\s.*items\b)\s{2,30}(\d{4})\s{2,30}\-?\+?\=?\s{2,30}(\(?(\d{1,3})?,?\d{3}\)?)?\s{2,30}\-?\+?\=?\s{2,30}(\(?(\d{1,3})?,?\d{3}\)?)?')
is_net_income = is_net_income_re.findall(is_text_p2)

