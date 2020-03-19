import openpyxl
from datetime import datetime

class StockManager:
    def __init__(self):
        self.workbook = None
        self.p_sheet = None
        self.t_sheet = None
        self.product_id = None
        self.type = None
        self.count = None


    def open_excel(self):
        self.workbook  = openpyxl.load_workbook('jiggingpro.xlsx')
        # self.t_sheet = self.workbook.get_sheet_by_name('Transaction')
        # self.p_sheet = self.workbook.get_sheet_by_name('Product')
        self.t_sheet = self.workbook['Transaction']
        self.p_sheet = self.workbook['Product']

    def _update_count(self, row, cnt):
        self.p_sheet['G' + str(row)] = self.p_sheet['G' + str(row)].value + cnt

    def _find_row_by_product_ID(self, id):
        ret = None
        for i in range(2, self.p_sheet.max_row+1):
            if self.p_sheet['A'+str(i)].value == id :
                ret = i
                break
            else :
                pass
        return ret


    def update_stock(self):
        get_product_ID = lambda row : self.t_sheet['C'+str(row)].value
        get_count = lambda row : self.t_sheet['D'+str(row)].value
        get_type = lambda row : self.t_sheet['B'+str(row)].value
        p_n = lambda t : 1 if t == 'buy' else -1

        for row in range(2, self.t_sheet.max_row+1):
            id = get_product_ID(row)
            type = get_type(row)
            cnt = get_count(row)
            cnt = cnt * p_n(type)

            p_row = self._find_row_by_product_ID(id)
            self._update_count(p_row, cnt)

    def save(self):
        now = datetime.now()
        fname = 'jiggingpro'+str(now.year)+'-'+str(now.month)+'-'+str(now.day)+'_'+str(now.hour)+str(now.minute)+str(now.second)+'.xlsx'
        self.workbook.save(fname)



sm = StockManager()
sm.open_excel()
sm.update_stock()
sm.save()



