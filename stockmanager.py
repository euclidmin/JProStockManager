# import pandas as pd
# import openpyxl
# import xlrd
#
# # df_transaction = pd.read_excel('jiggingpro.xlsx', sheet_name='Transaction')
# # df_price = pd.read_excel('jiggingpro.xlsx', sheet_name='Price')
# # df_product = pd.read_excel('jiggingpro.xlsx', sheet_name='Product')
# # df_manufacturer = pd.read_excel('jiggingpro.xlsx', sheet_name='Manufacturer')
# # df_distributor = pd.read_excel('jiggingpro.xlsx', sheet_name='Distributor')
# #
# # print(df_transaction)
# # print(df_price)
# # print(df_product)
# # print(df_manufacturer)
# # print(df_distributor)
# #
# # # excel 파일 읽어서 list 저장
# # wb = xlrd.open_workbook('jiggingpro.xlsx')
# # sh_tran = wb.sheet_by_name('Transaction')
# # sh_prod = wb.sheet_by_name('Product')
# #
# # # print(sh_prod)
# # # col0 = sh_prod.col(0)
# # ncol = sh_prod.ncols
# # nrow = sh_prod.nrows
# # # print(col0)
# #
# # # colomn 개수, row 개수
# # print(ncol)
# # print(nrow)
#
#
# def get_product(prod_id):
#     for idx in range(1, sh_prod.nrows):
#         rec_p = sh_prod.row(idx)
#         print('rec_p[0]= {0} prod_id={1}'.format(rec_p[0], prod_id))
#         if (rec_p[0].value == prod_id) :
#             ret = rec_p
#             # return ret
#             break
#         else :
#             ret = None
#     return ret
#
# def update_stock_cnt(rec_p, cnt, type):
#     if type == 'sell' : cnt *= -1
#     rec_p[6].value += cnt
#
#
# # def get_product_ID(rec_t):
# #     return rec_t[2].value
# get_product_ID = lambda rec_t : rec_t[2].value
#
# def get_transaction_cnt(rec_t):
#     return rec_t[3].value
#
# def get_transaction_type(rec_t):
#     return rec_t[1].value
#
#
#
#
#
#
#
# # excel 파일 읽어서 list 저장
# wb = xlrd.open_workbook('jiggingpro.xlsx')
# sh_tran = wb.sheet_by_name('Transaction')
# sh_prod = wb.sheet_by_name('Product')
#
# # print(sh_prod)
# # col0 = sh_prod.col(0)
# ncol = sh_prod.ncols
# nrow = sh_prod.nrows
# # print(col0)
#
# # colomn 개수, row 개수
# print(ncol)
# print(nrow)
#
#
# get_record = lambda idx : sh_tran.row(idx)
# #재고 계산
# # 트렌젝션 레코드에서 제품 레코드 찾아서 재고 카운트 업데이트 한다.
# for idx in range(1, sh_tran.nrows):
#     # rec_t = sh_tran.row(idx)
#     rec_t = get_record(idx)
#     print(rec_t)
#     prod_id = get_product_ID(rec_t)
#     rec_p = get_product(prod_id)                     # rec_t[2] =>  Product ID
#     cnt = get_transaction_cnt(rec_t)
#     t_type = get_transaction_type(rec_t)
#     update_stock_cnt(rec_p, cnt, t_type)       # rec_t[3] =>  Count , rec_t[1] => Type(sell/buy)
#     print(rec_p)
#     # 업데이트한 레코드를 저장 하지 않으면 사라짐,
#     # sheet 객체에 업데이트 하고 최종 파일에 업데이트 하는 방법
#     # 아니면 해달 셀만 파일에 직접 업데이트 하는 방법
#
#
# for idx in range(1, sh_prod.nrows):
#     # row 함수는 sheet 객체에서 List 객체를 생성 해서 반환 한다, 부를때 마다 새로 생성 한다
#     rec_p = sh_prod.row(idx)
#     print(rec_p)
#
#
#


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
        # get_product_ID = lambda row : self.sheet.cell(row, column_index_from_string('C')).value
        # get_count = lambda row : self.sheet.cell(row, column_index_from_string('D')).value
        # get_type = lambda row : self.sheet.cell(row, column_index_from_string('B')).value
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
        # fname = 'jiggingpro'+str(now.year)+'-'+str(now.month)+'-'+str(now.day)+'_'+str(now.hour)+':'+str(now.minute)+':'+str(now.second)+'.xlsx'
        fname = 'jiggingpro'+str(now.year)+'-'+str(now.month)+'-'+str(now.day)+'_'+str(now.hour)+str(now.minute)+str(now.second)+'.xlsx'
        self.workbook.save(fname)






sm = StockManager()
sm.open_excel()
sm.update_stock()
sm.save()



