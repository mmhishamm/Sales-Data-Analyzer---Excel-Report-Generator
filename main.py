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

TopProductSold = QuantityByProduct.sort_values(ascending=False).head(1)
TopRevenuePercentage = RevenuePercentage.sort_values(ascending=False).head(1)
TopRevenueByProduct = RevenueByProduct.sort_values(ascending=False).head(1)
TopAvgPriceByProduct = AvgPriceByProduct.sort_values(ascending=False).head(1)
TopProductCount = ProductCount.sort_values(ascending=False).head(1)

TopProducts = pd.DataFrame({"Metrics": ["Top Average Price", "Top Product Sold", "Top Revenue Product", "Top Revenue Percentage", "Top orders made"],
                            "Product":[TopAvgPriceByProduct.index[0],TopProductSold.index[0],TopRevenueByProduct.index[0],TopRevenuePercentage.index[0],TopProductCount.index[0]],
                             "Values":[TopAvgPriceByProduct.iloc[0],TopProductSold.iloc[0],TopRevenueByProduct.iloc[0],TopRevenuePercentage.iloc[0],TopProductCount.iloc[0]] })


TotalRevenue = MainExcelFile["Revenue"].sum()
TotalUnits = MainExcelFile["Quantity"].sum()
TotalOrders= MainExcelFile.index.size
AverageOrderValue = MainExcelFile["Revenue"].mean()
AverageUnitPrice = MainExcelFile["Unit_Price"].mean()

Summary = pd.DataFrame({"Metric":["TotalRevenue", "TotalUnits", "TotalOrders", "AverageOrderValue","AverageUnitPrice"],
                        "Values":[TotalRevenue,TotalUnits, TotalOrders, AverageOrderValue,AverageUnitPrice ]})

with pd.ExcelWriter("Sales Report.xlsx", engine="openpyxl") as writer:
    Summary.to_excel(writer, sheet_name="Sales_Summary", index= False)
    ProductPerformance.to_excel(writer, sheet_name="Product_Performance")
    TopProducts.to_excel(writer, sheet_name="Product_Performance", startrow= len(ProductPerformance) + 3, index = False)