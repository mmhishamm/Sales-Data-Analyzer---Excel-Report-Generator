import pandas as pd

MainExcelFile = pd.read_excel("company_sales.xlsx")

print(MainExcelFile.to_string())