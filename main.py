import pandas as pd

MainExcelFile = pd.read_excel("company_sales.xlsx")

MainExcelFile = MainExcelFile.drop_duplicates()
MainExcelFile["City"] = MainExcelFile["City"].fillna("Unknown")
MainExcelFile["Unit_Price"] = MainExcelFile["Unit_Price"].fillna(MainExcelFile[MainExcelFile["Product"] == "Laptop"]["Unit_Price"].median())
MainExcelFile.loc[MainExcelFile["Quantity"] == 99, "Quantity"] = 9

