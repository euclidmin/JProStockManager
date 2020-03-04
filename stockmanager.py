import pandas as pd

df_transaction = pd.read_excel('jiggingpro.xlsx', sheet_name='Transaction')
df_price = pd.read_excel('jiggingpro.xlsx', sheet_name='Price')
df_product = pd.read_excel('jiggingpro.xlsx', sheet_name='Product')
df_manufacturer = pd.read_excel('jiggingpro.xlsx', sheet_name='Manufacturer')
df_distributor = pd.read_excel('jiggingpro.xlsx', sheet_name='Distributor')


print(df_transaction)
print(df_price)
print(df_product)
print(df_manufacturer)
print(df_distributor)


