import unittest
import os
from scripts import xlsx_writer


class Test(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        print(f"Starting tests for XLSX WRITER")

    @classmethod
    def tearDownClass(cls):
        print(f"Finishing tests for XLSX WRITER")

    def test_load_template(self):
        xlsx_writer.load_template()

    def test_save_template(self):
        template_file, template_sheet = xlsx_writer.load_template()
        xlsx_writer.save_file(template_file)
        template_route = xlsx_writer.obtain_template_location()
        if os.path.exists(f"{template_route}//written-excel.xlsx"):
            os.remove(f"{template_route}//written-excel.xlsx")

