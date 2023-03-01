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
def start(balance_sheet, income_statement_sheet):
    # balance_sheet = 'temp/balance_sheet2.pdf'
    # income_statement_sheet = 'temp/income_statement2.pdf'
    # get output name from the balance sheet name
    output_name = balance_sheet.replace('-S125-', '')
    output_name = output_name.replace('-S100-', '')
    output_name = output_name.replace('.pdf', '')
    output_name = re.sub("\d", '', output_name)

    output_file = 'output/' + output_name + '.xlsx'

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

    # ============================================ create excel file ==================================
    workbook = xlsxwriter.Workbook(output_file)

    # ///////////////////////// styles ///////////////////////////

    # Create a format to use in the merged range.
    initial_format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 12,
        'font_name': 'Arial',
    })
    title_top_left = copy_format(workbook, initial_format)
    title_top_left.set_top(2)
    title_top_left.set_align('left')

    font10 = copy_format(workbook, initial_format)
    font10.set_font_size(10)

    font10_left = copy_format(workbook, initial_format)
    font10_left.set_align('left')
    font10_left.set_font_size(10)

    font10_left_no_bold = copy_format(workbook, initial_format)
    font10_left_no_bold.set_align('left')
    font10_left_no_bold.set_font_size(10)
    font10_left_no_bold.set_bold(False)

    font10_top_left_bold = copy_format(workbook, initial_format)
    font10_top_left_bold.set_align('left')
    font10_top_left_bold.set_top(1)
    font10_top_left_bold.set_font_size(10)
    font10_top_left_bold.set_bold(True)

    bottom10 = copy_format(workbook, initial_format)
    bottom10.set_bottom(2)
    bottom10.set_font_size(10)

    bottom_left = copy_format(workbook, initial_format)
    bottom_left.set_bottom(2)
    bottom_left.set_align('left')
    bottom_left.set_font_size(10)

    bottom_one = copy_format(workbook, initial_format)
    bottom_one.set_bottom(1)

    section_title = copy_format(workbook, initial_format)
    section_title.set_align('left')
    section_title.set_underline(1)
    section_title.set_font_size(10)

    currency_format = workbook.add_format({'num_format': '_(\$* #,##0_);_(\$* (#,##0);_(\$* -_);_(@_)'})
    currency_format.set_bold(True)

    currency_top = workbook.add_format({'num_format': '_(\$* #,##0_);_(\$* (#,##0);_(\$* -_);_(@_)'})
    currency_top.set_bold(True)
    currency_top.set_top(1)

    currency_top_bottom = workbook.add_format({'num_format': '_(\$* #,##0_);_(\$* (#,##0);_(\$* -_);_(@_)'})
    currency_top_bottom.set_bold(True)
    currency_top_bottom.set_top(1)
    currency_top_bottom.set_bottom(6)

    currency_bottom_double = workbook.add_format({'num_format': '_(\$* #,##0_);_(\$* (#,##0);_(\$* -_);_(@_)'})
    currency_bottom_double.set_bold(True)
    currency_bottom_double.set_bottom(6)

    # ////////////////////////////// Title Sheet ////////////////////////////////
    title_sheet = workbook.add_worksheet('Title')

    # Increase the cell size of the merged cells to highlight the formatting.
    title_sheet.set_row(1, 20)
    title_sheet.set_row(2, 20)
    title_sheet.set_row(3, 20)

    title_sheet.merge_range('A2:G2', company_name, initial_format)
    title_sheet.merge_range('A3:G3', 'FINANCIAL STATEMENTS', initial_format)
    title_sheet.merge_range('A4:G4', 'AS AT ' + final_date, initial_format)

    # /////////////////////////////////// Balance Sheet ///////////////////////
    balance_sheet = workbook.add_worksheet('Balance Sheet')

    balance_sheet.set_row(1, 20)
    balance_sheet.set_column('B:B', 50)
    balance_sheet.set_column('C:C', 10)
    balance_sheet.set_column('D:D', 3)
    balance_sheet.set_column('F:F', 3)
    balance_sheet.set_column('E:E', 10)
    balance_sheet.set_column('G:G', 10)

    # TOP TITLE
    balance_sheet.merge_range('A2:G2', company_name, title_top_left)
    balance_sheet.merge_range('A3:G3', 'BALANCE SHEET', font10_left)
    balance_sheet.merge_range('A4:G4', 'AS AT ' + final_date, font10_left)
    balance_sheet.merge_range('A5:G5', '(UNAUDITED - SEE NOTICE TO READER)', bottom_left)

    # ASSETS
    balance_sheet.write('A8', 'ASSETS:', section_title)
    balance_sheet.write('E8', int(year), bottom10)
    balance_sheet.write('G8', ( int(year) - 1 ), bottom10)

    balance_sheet.write('B10', 'CURRENT ASSETS:', font10_left)
    balance_sheet.write('C11', '(Notes)', font10_left_no_bold)

    row_num = 12
    for x in assets:
        title =  x[0].rstrip()
        if title == 'Total assets':
            row_num += 1
            balance_sheet.write('B' + str(row_num), title, font10_top_left_bold)
            balance_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            balance_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_top)
        else:
            balance_sheet.write('B' + str(row_num), title, font10_left_no_bold)
            balance_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            balance_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_format)
        row_num += 1
        
    # LIABILITIES 
    row_num += 2
    balance_sheet.write('A' + str(row_num), 'LIABILITIES & SHAREHOLDER\'S EQUITY:', section_title)
    row_num += 2
    balance_sheet.write('B' + str(row_num), 'CURRENT LIABILITIES:', font10_left)
    row_num += 2

    for x in liabilities:
        # print(clean_amount(x[2]))
        title =  x[0].rstrip()
        if title == 'Total liabilities':
            row_num += 1
            balance_sheet.write('B' + str(row_num), 'LONG TERM DEBTS:', font10_left)
            row_num += 2
            balance_sheet.write('B' + str(row_num), 'Long Term Debt', font10_left_no_bold)
            balance_sheet.write('E' + str(row_num), 0, currency_format)
            balance_sheet.write('G' + str(row_num), 0, currency_format)
            row_num += 1
            balance_sheet.write('B' + str(row_num), 'Due to Shareholder', font10_left_no_bold)
            balance_sheet.write('E' + str(row_num), 0, currency_format)
            balance_sheet.write('G' + str(row_num), 0, currency_format)
            row_num += 1
            balance_sheet.write('B' + str(row_num), title, font10_top_left_bold)
            balance_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            balance_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            balance_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_top)
        else:
            balance_sheet.write('B' + str(row_num), title, font10_left_no_bold)
            balance_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            balance_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_format)
        row_num += 1

    # Equity 
    row_num += 2
    balance_sheet.write('B' + str(row_num), 'SHAREHOLDERS EQUITY:', font10_left)
    row_num += 2
    balance_sheet.write('B' + str(row_num), 'Retained Earnings', font10_left_no_bold)
    balance_sheet.write('E' + str(row_num), clean_amount(equities[0][2]), currency_format)
    balance_sheet.write('G' + str(row_num), clean_amount(equities[0][4]), currency_format)
    row_num += 1
    balance_sheet.write('B' + str(row_num), ' ', font10_top_left_bold)
    balance_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
    balance_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
    balance_sheet.write('E' + str(row_num), clean_amount(equities[0][2]), currency_top)
    balance_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
    balance_sheet.write('G' + str(row_num), clean_amount(equities[0][4]), currency_top)
    row_num += 2
    balance_sheet.write('A' + str(row_num), 'TOTAL LIABILITIES & SHAREHOLDERS EQUITY', font10_left) 
    balance_sheet.write('E' + str(row_num), clean_amount(equities[2][2]), currency_top_bottom)
    balance_sheet.write('G' + str(row_num), clean_amount(equities[2][4]), currency_top_bottom)
    row_num += 1
    balance_sheet.write('E' + str(row_num), ' ', currency_top)
    balance_sheet.write('G' + str(row_num), ' ', currency_top)
    row_num += 2
    balance_sheet.write('A' + str(row_num), 'Director:', font10_left_no_bold) 
    balance_sheet.write('B' + str(row_num), ' ', bottom_one) 

    # ////////////////////////////////////// (INCOME STATEMENT) IS Sheet ///////////////////////
    is_sheet = workbook.add_worksheet('IS')
    row_num = 0

    is_sheet.set_row(1, 20)
    is_sheet.set_column('B:B', 50)
    is_sheet.set_column('C:C', 10)
    is_sheet.set_column('D:D', 3)
    is_sheet.set_column('F:F', 3)
    is_sheet.set_column('E:E', 10)
    is_sheet.set_column('G:G', 10)

    # TOP TITLE
    is_sheet.merge_range('A2:G2', company_name, title_top_left)
    is_sheet.merge_range('A3:G3', 'INCOME STATEMENT', font10_left)
    is_sheet.merge_range('A4:G4', 'AS AT ' + final_date, font10_left)
    is_sheet.merge_range('A5:G5', '(UNAUDITED - SEE NOTICE TO READER)', bottom_left)
    row_num += 7

    # REVENUE
    is_sheet.write('A' + str(row_num), 'REVENUE:', font10_left)
    is_sheet.write('E' + str(row_num), int(year), bottom10)
    is_sheet.write('G' + str(row_num), ( int(year) - 1 ), bottom10)
    row_num += 2
    is_sheet.write('B' + str(row_num), 'Revenue', font10_left_no_bold)
    is_sheet.write('E' + str(row_num), clean_amount(is_revenue[2][2]), currency_format)
    is_sheet.write('G' + str(row_num), clean_amount(is_revenue[2][5]), currency_format)
    row_num += 1
    is_sheet.write('B' + str(row_num), ' ', font10_top_left_bold)
    is_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
    is_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
    is_sheet.write('E' + str(row_num), ' ', currency_top)
    is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
    is_sheet.write('G' + str(row_num), ' ', currency_top)
    row_num += 1

    # Cost Of Sales
    is_sheet.write('B' + str(row_num), 'COST OF SALES:', font10_left)
    row_num += 2

    for x in is_cos:
        title =  x[0].rstrip()
        if title == 'Cost  of  sales':
            is_sheet.write('B' + str(row_num), 'Cost of sales', font10_top_left_bold)
            is_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_top)
        elif 'Gross   profit' in title:
            row_num += 1
            is_sheet.write('A' + str(row_num), 'GROSS MARGIN', font10_left)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_top)
        else:
            is_sheet.write('B' + str(row_num), title, font10_left_no_bold)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_format)
        row_num += 1

    # Operating expenses
    row_num += 1
    is_sheet.write('B' + str(row_num), 'ADMINISTRATIVE/OPERATING EXPENSES:', font10_left)
    row_num += 2
    for x in is_oe:
        title =  x[0].rstrip()
        if title == 'Total  operating   expenses':
            row_num += 1
            is_sheet.write('B' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('E' + str(row_num), ' ', currency_top)
            is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('G' + str(row_num), ' ', currency_top)
            row_num += 1
            is_sheet.write('B' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_top)
        elif 'Net  non-farming     income' in title:
            row_num += 1
            is_sheet.write('B' + str(row_num), 'INCOME / (LOSS) FROM OPERATION', font10_left)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_top)
            is_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_top)
        elif 'Total  expenses' in title:
            continue
        else:
            is_sheet.write('B' + str(row_num), title, font10_left_no_bold)
            is_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            is_sheet.write('G' + str(row_num), clean_amount(x[5]), currency_format)
        row_num += 1
    row_num += 1
    is_sheet.write('B' + str(row_num), 'Current Income Taxes', font10_left)
    is_sheet.write('E' + str(row_num), clean_amount(is_current_taxes[0][2]), currency_format)
    is_sheet.write('G' + str(row_num), clean_amount(is_current_taxes[0][4]), currency_format)
    row_num += 2
    is_sheet.write('B' + str(row_num), 'INCOME / (LOSS) AFTER TAXES', font10_left)
    is_sheet.write('E' + str(row_num), clean_amount(is_net_income[0][2]), currency_format)
    is_sheet.write('G' + str(row_num), clean_amount(is_net_income[0][4]), currency_format)

    # ///////////////////////////////// STATEMENT OF RETAINED EARNING / (DEFICIT) RE Sheet ///////////////////////
    re_sheet = workbook.add_worksheet('RE')
    row_num = 0

    re_sheet.set_row(1, 20)
    re_sheet.set_column('B:B', 50)
    re_sheet.set_column('C:C', 10)
    re_sheet.set_column('D:D', 3)
    re_sheet.set_column('F:F', 3)
    re_sheet.set_column('E:E', 10)
    re_sheet.set_column('G:G', 10)

    # TOP TITLE
    re_sheet.merge_range('A2:G2', company_name, title_top_left)
    re_sheet.merge_range('A3:G3', 'STATEMENT OF RETAINED EARNING / (DEFICIT)', font10_left)
    re_sheet.merge_range('A4:G4', 'AS AT ' + final_date, font10_left)
    re_sheet.merge_range('A5:G5', '(UNAUDITED - SEE NOTICE TO READER)', bottom_left)
    row_num += 7

    re_sheet.write('E' + str(row_num), int(year), bottom10)
    re_sheet.write('G' + str(row_num), ( int(year) - 1 ), bottom10)
    row_num += 3

    re_sheet.write('A' + str(row_num), 'STATEMENT OF RETAINED EARNINGS', section_title)
    row_num += 2

    for x in retained_earnings:
        title =  x[0].rstrip()
        if 'deficit-start' in title:
            re_sheet.write('B' + str(row_num), 'Balance at beginning of year', font10_left_no_bold)
            re_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            re_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_format)
            row_num += 1
        elif 'Dividends declared' in title:
            re_sheet.write('B' + str(row_num), 'Dividend Issued', font10_left_no_bold)
            re_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            re_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_format)
            row_num += 1
    for x in retained_earnings:
        title =  x[0].rstrip()
        if 'Net income / loss' in title:
            re_sheet.write('B' + str(row_num), 'Net Income / (Loss) for the year', font10_left_no_bold)
            re_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_format)
            re_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_format)
            row_num += 1
            re_sheet.write('B' + str(row_num), ' ', font10_top_left_bold)
            re_sheet.write('C' + str(row_num), ' ', font10_top_left_bold)
            re_sheet.write('D' + str(row_num), ' ', font10_top_left_bold)
            re_sheet.write('E' + str(row_num), ' ', currency_top)
            re_sheet.write('F' + str(row_num), ' ', font10_top_left_bold)
            re_sheet.write('G' + str(row_num), ' ', currency_top)
            row_num += 1
        elif 'Total retained earnings' in title:
            re_sheet.write('B' + str(row_num), 'Retained earning / (deficit) at the end of year', font10_left_no_bold)
            re_sheet.write('E' + str(row_num), clean_amount(x[2]), currency_bottom_double)
            re_sheet.write('G' + str(row_num), clean_amount(x[4]), currency_bottom_double)

    # close the workbook
    workbook.close()
    
    return output_name