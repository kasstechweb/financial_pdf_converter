import re
import pdfplumber
import pandas as pd
from datetime import datetime
import fitz
import os
from PIL import Image
import pytesseract
import xlsxwriter

balance_sheet = 'temp/balance_sheet2.pdf'
income_statement_sheet = 'temp/income_statement2.pdf'
assets = []
liabilities = []
equities = []
retained_earnings = []

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
