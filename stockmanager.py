import openpyxl
from datetime import datetime
from PySide2 import QtCore, QtWidgets, QtGui
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile
import sys
from PySide2.QtWidgets import QMainWindow, QWidget


class StockManager:
    def __init__(self):
        self.workbook = None
        self.p_sheet = None
        self.t_sheet = None
        self.m_sheet = None
        self.product_id = None
        self.type = None
        self.count = None


    def open_excel(self):
        self.workbook  = openpyxl.load_workbook('jiggingpro.xlsx')
        # self.t_sheet = self.workbook.get_sheet_by_name('Transaction')
        # self.p_sheet = self.workbook.get_sheet_by_name('Product')
        self.t_sheet = self.workbook['Transaction']
        self.p_sheet = self.workbook['Product']
        self.m_sheet = self.workbook['Manufacturer']


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

    def get_manufacturer(self):
        manuf_cells = lambda colm: self.m_sheet[colm]
        cell_value = lambda cell : cell.value

        manufacturer_list = list(map(cell_value, manuf_cells('B')))
        return manufacturer_list[1:]         # 첫줄의 카테고리 이름을 제외 시킨다.











class stock_window:
    def __init__(self, stock_manager):
        self.sm = stock_manager
        self.ui_file = QFile("product_select.ui")
        self.ui_file.open(QFile.ReadOnly)
        self.loader = QUiLoader()
        self.main_window = self.loader.load(self.ui_file)
        self.ui_file.close()
        self.main_window.show()

        self.combobox_add_manufacturer()

    def combobox_add_manufacturer(self):
        cb = self.main_window.findChild(QtWidgets.QComboBox, "comboBox")
        mf_list = self.sm.get_manufacturer()
        cb.addItems(mf_list)


if __name__ == '__main__':
    sm = StockManager()
    sm.open_excel()
    # sm.update_stock()
    # sm.save()
    app = QtWidgets.QApplication(sys.argv)
    sw = stock_window(sm)
    sys.exit(app.exec_())
    

