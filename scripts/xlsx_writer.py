import openpyxl
import os

TEMPLATE_NAME = "template.xlsx"
OUTPUT_NAME = "written-excel.xlsx"


def obtain_template_location():
    actual_route = os.path.dirname(__file__)
    template_route = actual_route + f"\\..\\templates\\"
    return template_route


def load_template():
    template_route = obtain_template_location()
    template_file = openpyxl.load_workbook(filename=template_route + TEMPLATE_NAME)
    template_sheet = template_file[template_file.sheetnames[0]]
    return template_file, template_sheet


def save_file(file):
    template_route = obtain_template_location()
    file.save(template_route + OUTPUT_NAME)


def write_value_in_determined_table_row_column(value, table, row, column):
    table[f"{column}{row}"] = value
