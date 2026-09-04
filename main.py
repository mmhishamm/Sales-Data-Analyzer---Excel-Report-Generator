import pandas as pd

MainExcelFile = pd.read_excel("company_sales.xlsx", index_col="Order_ID")

MainExcelFile = MainExcelFile.drop_duplicates()
MainExcelFile["City"] = MainExcelFile["City"].fillna("Unknown")
MainExcelFile["Unit_Price"] = MainExcelFile["Unit_Price"].fillna(MainExcelFile[MainExcelFile["Product"] == "Laptop"]["Unit_Price"].median())
MainExcelFile.loc[MainExcelFile["Quantity"] == 99, "Quantity"] = 9
MainExcelFile["Revenue"] = MainExcelFile["Quantity"] * MainExcelFile ["Unit_Price"]
MainExcelFile.to_excel("clean_data.xlsx")

AvgPriceByProduct = MainExcelFile.groupby("Product")["Unit_Price"].mean()
AvgPriceByProduct.name = "Average Unit Price"
QuantityByProduct = MainExcelFile.groupby("Product")["Quantity"].sum()
QuantityByProduct.name = "Quantity"
RevenueByProduct = MainExcelFile.groupby("Product")["Revenue"].sum()
RevenueByProduct.name = "Revenue"
ProductCount = MainExcelFile.groupby("Product").size()
ProductCount.name = "Product count"
RevenuePercentage = (MainExcelFile.groupby("Product")["Revenue"].sum() / MainExcelFile["Revenue"].sum()) * 100
RevenuePercentage.name = "Revenue Percentage"

ProductPerformance = pd.concat([AvgPriceByProduct,
                                QuantityByProduct,
                                  RevenueByProduct, 
                                  ProductCount,
                                  RevenuePercentage], axis=1)

ProductPerformance.loc["Total"] = [MainExcelFile["Unit_Price"].mean(),
                                    MainExcelFile["Quantity"].sum(),
                                      MainExcelFile["Revenue"].sum(),
                                      MainExcelFile["Product"].size,
                                      100,
                                      ] 

ProductPerformance.to_excel("Product_Performance.xlsx")