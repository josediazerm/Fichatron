import openpyxl
import os

template_file = openpyxl.load_workbook(filename=f'{os.path.dirname(__file__)}\\template.xlsx')
template_sheet = template_file[template_file.sheetnames[0]]
