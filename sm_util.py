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
            for col in range(self.sheet.ncols):
                self.cell[row][col] = self.sheet.cell_value(row, col)


        






exlm = ExcelManager()
exlm.open_workbook('jiggingpro.xlsx')
exlm.get_sheet('Transaction')
exlm.load_data()

exlm.workbook =
# ws_product = wb.sheet_by_name('Product')





