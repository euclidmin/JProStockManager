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


def get_product(prod_id):
    for idx in range(1, sh_prod.nrows):
        rec_p = sh_prod.row(idx)
        print('rec_p[0]= {0} prod_id={1}'.format(rec_p[0], prod_id))
        if (rec_p[0].value == prod_id) :
            ret = rec_p
            return ret
        else :
            ret = None
    return ret

def update_stock_cnt(rec_p, cnt, type):
    if type == 'sell' : cnt *= -1
    rec_p[6].value += cnt


def get_product_ID(rec_t):
    return rec_t[2].value

def get_transaction_cnt(rec_t):
    return rec_t[3].value

def get_transaction_type(rec_t):
    return rec_t[1].value


#재고 계산
# 트렌젝션 레코드에서 제품 레코드 찾아서 재고 카운트 업데이트 한다.
for idx in range(1, sh_tran.nrows):
    rec_t = sh_tran.row(idx)
    print(rec_t)
    prod_id = get_product_ID(rec_t)
    rec_p = get_product(prod_id)                     # rec_t[2] =>  Product ID
    cnt = get_transaction_cnt(rec_t)
    t_type = get_transaction_type(rec_t)
    update_stock_cnt(rec_p, cnt, t_type)       # rec_t[3] =>  Count , rec_t[1] => Type(sell/buy)

#
# for idx in range(1, sh_prod.nrows):
#     rec_p = sh_prod.row(idx)
#     print(rec_p)













