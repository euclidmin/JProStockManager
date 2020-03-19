import xlrd, xlwt


class ExcelManager:
    def __init__(self):
        self.workbook = None
        self.sheet = None
        self.cell = {}
#         초기화 필요한 것들 잡아 준다.

    def open_workbook(self, name):
        self.workbook = xlrd.open_workbook(name)

    def get_sheet(self, name):
        self.sheet = self.workbook.sheet_by_name(name)

    def load_data(self):
        for row in range(self.sheet.nrows):
            self.cell[row] = {}
            for col in range(self.sheet.ncols):
                self.cell[row][col] = self.sheet.cell_value(row, col)

    def get_cell_data(self, row, col):
        return self.cell[row][col]

    def get_row(self, row):
        return self.cell[row]

    def set_cell_data(self, row, col, value):
        self.cell[row][col] = value

    def print_sheet(self):
        for i in range(self.sheet.nrows):
            print(self.get_row(i))





exlm = ExcelManager()
exlm.open_workbook('jiggingpro.xlsx')
exlm.get_sheet('Transaction')
exlm.load_data()

exlm.print_sheet()
exlm.set_cell_data(0, 0, 'id')
exlm.print_sheet()




