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
        test_passed = False
        if os.path.exists(f"{template_route}//written-excel.xlsx"):
            test_passed = True
            os.remove(f"{template_route}//written-excel.xlsx")
        self.assertTrue(test_passed)

    def test_write_value(self):
        template_file, template_sheet = xlsx_writer.load_template()
        test_value = 10
        xlsx_writer.write_value_in_determined_table_row_column(value=test_value,
                                                               table=template_sheet,
                                                               row=1,
                                                               column="A")
        self.assertEqual(template_sheet["A1"].value, test_value)

