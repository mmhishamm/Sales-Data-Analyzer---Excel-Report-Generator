import pandas as pd

MainExcelFile = pd.read_excel("company_sales.xlsx")

print(MainExcelFile.head())
print(MainExcelFile.shape)
print(MainExcelFile.columns)
print(MainExcelFile.dtypes)
print(MainExcelFile.isna().sum())
print(MainExcelFile.duplicated().sum())