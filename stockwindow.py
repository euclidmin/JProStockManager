import sys
from PySide2 import QtWidgets
from PySide2.QtCore import QFile
from PySide2.QtUiTools import QUiLoader
from excelmanager import ExcelManager


class StockWindow:

    def __init__(self, excel_manager):
        self.excel_manager = excel_manager
        self.ui_file = QFile("product_select.ui")
        self.ui_file.open(QFile.ReadOnly)
        self.loader = QUiLoader()
        self.main_window = self.loader.load(self.ui_file)
        self.ui_file.close()
        self.main_window.show()

        # 콤보박스 객체 저장
        self.combobox_manufacturer = None
        self.combobox_product = None
        self.combobox_weight = None
        self.combobox_color = None
        self.combobox_size = None

        # 콤보박스 선택된 item Text
        self.combobox_manufacturer_text = None
        self.combobox_product_text = None
        self.combobox_weight_text = None
        self.combobox_color_text = None
        self.combobox_size_text = None

        # 제조사 콤보박스 객체 찾기
        cb = self.main_window.findChild(QtWidgets.QComboBox, "comboBox")
        self.combobox_manufacturer = cb
        # 제조사 콤보박스에 시그널 별로 슬랏 연결
        cb.activated.connect(self.activated_manufacturer)

        # 제품 콤보박스 객체 찾기
        cb2 = self.main_window.findChild(QtWidgets.QComboBox, "comboBox_2")
        self.combobox_product = cb2
        # 제품 콤보박스에 시그널 슬랏 연결
        cb2.activated.connect(self.activated_product)

        # 무게 콤보박스 객체 찾기
        cb3 = self.main_window.findChild(QtWidgets.QComboBox, "comboBox_3")
        self.combobox_weight = cb3
        # 제품 콤보박스에 시그널 슬랏 연결
        cb3.activated.connect(self.activated_weight)

        # 컬러 콤보박스 객체 찾기
        cb4 = self.main_window.findChild(QtWidgets.QComboBox, "comboBox_4")
        self.combobox_color = cb4
        # 제품 콤보박스에 시그널 슬랏 연결
        cb3.activated.connect(self.activated_color)

        self.combobox_add_manufacturer()



    def combobox_add_manufacturer(self):
        cb = self.combobox_manufacturer
        # 엑셀 메니저에서 제조사 리스트 가져오기
        mf_list = self.excel_manager.get_manufacturer()
        # 제조사 리스트 콤보박스에 아이템 추가
        cb.addItems(mf_list)



    def activated_manufacturer(self):
        print("activated")
        # 현재 선택된 아이템 저장
        text = self.combobox_manufacturer.currentText()
        self.combobox_manufacturer_text = text
        print(text)

        # 선택된 제조사를 보고 그 제조사 제품을 추가
        self.combobox_add_product()

    def combobox_add_product(self):
        # 제조사에 해당하는 제품 리스트 가져오기
        name = self.combobox_manufacturer_text
        p_list = self.excel_manager.get_products(name)
        print(p_list)

        # 제품 콤보박스에 제품리스트 추가
        cb2 = self.combobox_product
        cb2.addItems(p_list)

    def activated_product(self):
        print("activated")
        # 현재 선택된 아이템 저장
        text = self.combobox_product.currentText()
        self.combobox_product_text = text
        print(text)

        # 선택된 제조사를 보고 그 제조사 제품을 추가
        self.combobox_add_weight()

    def combobox_add_weight(self):
        name = self.combobox_product_text
        w_list = self.excel_manager.get_weight(name)
        print(w_list)

        cb3 = self.combobox_weight
        cb3.addItems(w_list)


    def activated_weight(self):
        print("activated")
        # 현재 선택된 아이템 저장
        text = self.combobox_weight.currentText()
        self.combobox_weight_text = text
        print(text)

        # 선택된 제조사를 보고 그 제조사 제품을 추가
        self.combobox_add_color()

    def combobox_add_color(self):
        name = self.combobox_weight_text
        c_list = self.excel_manager.get_color(name)
        print(c_list)

        cb4 = self.combobox_color
        cb4.addItems(c_list)

    def activated_color(self):
        print("activated")
        # 현재 선택된 아이템 저장
        text = self.combobox_color.currentText()
        self.combobox_color_text = text
        print(text)

        # 선택된 제조사를 보고 그 제조사 제품을 추가
        # self.combobox_add_size()

    def highlighted(self):
        print("highlighted")

    def idxchanged(self):
        print("idx changed")


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    sm = ExcelManager()
    sm.open_excel()
    sw = StockWindow(sm)
    # sm.update_stock()
    # sm.save()
    sys.exit(app.exec_())

