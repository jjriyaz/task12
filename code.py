from openpyexcel import load_workbook


class all_excel_functions:
    def __init__(self, excel_file_name, sheet_name):
        self.file = excel_file_name
        self.sheet = sheet_name

    def row_count(self):
        workbook = load_workbook(self.file)
        sheet = workbook[self.sheet]
        return (sheet.max_row)

    def column_count(self):
        workbook = load_workbook(self.file)
        sheet = workbook[self.sheet]
        return (sheet.max_column)

    def read_data(self, row_number, column_number):
        workbook = load_workbook(self.file)
        sheet = workbook[self.sheet]
        data = sheet.cell(row=row_number, column=column_number).value
        return data

    def write_data(self, row_number, column_number, data):
        workbook = load_workbook(self.file)
        sheet = workbook[self.sheet]
        sheet.cell(row=row_number, column=column_number).value = data
        workbook.save(self.file)

