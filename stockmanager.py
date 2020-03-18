import pandas as pd
import openpyxl
import xlrd

# df_transaction = pd.read_excel('jiggingpro.xlsx', sheet_name='Transaction')
# df_price = pd.read_excel('jiggingpro.xlsx', sheet_name='Price')
# df_product = pd.read_excel('jiggingpro.xlsx', sheet_name='Product')
# df_manufacturer = pd.read_excel('jiggingpro.xlsx', sheet_name='Manufacturer')
# df_distributor = pd.read_excel('jiggingpro.xlsx', sheet_name='Distributor')
#
# print(df_transaction)
# print(df_price)
# print(df_product)
# print(df_manufacturer)
# print(df_distributor)
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


def get_product(prod_id):
    for idx in range(1, sh_prod.nrows):
        rec_p = sh_prod.row(idx)
        print('rec_p[0]= {0} prod_id={1}'.format(rec_p[0], prod_id))
        if (rec_p[0].value == prod_id) :
            ret = rec_p
            # return ret
            break
        else :
            ret = None
    return ret

def update_stock_cnt(rec_p, cnt, type):
    if type == 'sell' : cnt *= -1
    rec_p[6].value += cnt


# def get_product_ID(rec_t):
#     return rec_t[2].value
get_product_ID = lambda rec_t : rec_t[2].value

def get_transaction_cnt(rec_t):
    return rec_t[3].value

def get_transaction_type(rec_t):
    return rec_t[1].value







# excel 파일 읽어서 list 저장
wb = xlrd.open_workbook('jiggingpro.xlsx')
sh_tran = wb.sheet_by_name('Transaction')
sh_prod = wb.sheet_by_name('Product')

# print(sh_prod)
# col0 = sh_prod.col(0)
ncol = sh_prod.ncols
nrow = sh_prod.nrows
# print(col0)

# colomn 개수, row 개수
print(ncol)
print(nrow)


get_record = lambda idx : sh_tran.row(idx)
#재고 계산
# 트렌젝션 레코드에서 제품 레코드 찾아서 재고 카운트 업데이트 한다.
for idx in range(1, sh_tran.nrows):
    # rec_t = sh_tran.row(idx)
    rec_t = get_record(idx)
    print(rec_t)
    prod_id = get_product_ID(rec_t)
    rec_p = get_product(prod_id)                     # rec_t[2] =>  Product ID
    cnt = get_transaction_cnt(rec_t)
    t_type = get_transaction_type(rec_t)
    update_stock_cnt(rec_p, cnt, t_type)       # rec_t[3] =>  Count , rec_t[1] => Type(sell/buy)
    print(rec_p)
    # 업데이트한 레코드를 저장 하지 않으면 사라짐,
    # sheet 객체에 업데이트 하고 최종 파일에 업데이트 하는 방법
    # 아니면 해달 셀만 파일에 직접 업데이트 하는 방법


for idx in range(1, sh_prod.nrows):
    # row 함수는 sheet 객체에서 List 객체를 생성 해서 반환 한다, 부를때 마다 새로 생성 한다
    rec_p = sh_prod.row(idx)
    print(rec_p)













