import re
import os

def copy_format(book, fmt):
    properties = [f[4:] for f in dir(fmt) if f[0:4] == 'set_']
    dft_fmt = book.add_format()
    return book.add_format({k : v for k, v in fmt.__dict__.items() if k in properties and dft_fmt.__dict__[k] != v})

def clean_amount(str):
    if str == '':
        return 0
    else:
        new_str = str
        if '(' in new_str:
            new_str = new_str.replace('(', '')
            new_str = new_str.replace(')', '')
            new_str = new_str.replace(',', '')
            return_int = int(new_str) * -1
        elif new_str == '':
            return_int = 0
        else:
            new_str = new_str.replace(',', '')
            return_int = int(new_str)
        return return_int

def generate_output_name(balance_sheet_name):
    output_name = balance_sheet_name.replace('-S125-', '')
    output_name = output_name.replace('-S100-', '')
    output_name = output_name.replace('.pdf', '')
    output_name = output_name.replace('(', '')
    output_name = output_name.replace(')', '')
    output_name = re.sub("\d", '', output_name)

    # if directory has same file add numbers
    path = 'output/' + output_name + '.xlsx'
    num = 2
    new_output_name = ''
    while os.path.isfile(path):
        new_output_name = output_name + str(num)
        path = 'output/' + new_output_name + '.xlsx'
        num += 1
    # print(new_output_name)
    if new_output_name == '':
        new_output_name = output_name
    return new_output_name