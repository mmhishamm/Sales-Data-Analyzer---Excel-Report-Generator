import pandas as pd

def DataCleaning():
  MainExcelFile = pd.read_excel("company_sales.xlsx", index_col="Order_ID")
  MainExcelFile = MainExcelFile.drop_duplicates()
  MainExcelFile["City"] = MainExcelFile["City"].fillna("Unknown")
  MainExcelFile["Unit_Price"] = MainExcelFile["Unit_Price"].fillna(MainExcelFile[MainExcelFile["Product"] == "Laptop"]["Unit_Price"].median())
  MainExcelFile.loc[MainExcelFile["Quantity"] == 99, "Quantity"] = 9
  MainExcelFile["Revenue"] = MainExcelFile["Quantity"] * MainExcelFile ["Unit_Price"]
  MainExcelFile.to_excel("clean_data.xlsx")
  MainExcelFile["Month"]=MainExcelFile["Date"].dt.to_period("M").astype("str")
  return MainExcelFile

def Analysis(MainExcelFile,Metric, writer):
  AvgPrice = MainExcelFile.groupby(Metric)["Unit_Price"].mean()
  AvgPrice.name = "Average Unit Price"
  Quantity = MainExcelFile.groupby(Metric)["Quantity"].sum()
  Quantity.name = "Quantity"
  Revenue = MainExcelFile.groupby(Metric)["Revenue"].sum()
  Revenue.name = "Revenue"
  Count = MainExcelFile.groupby(Metric).size()
  Count.name = "Orders Made"
  RevenuePercentage = (MainExcelFile.groupby(Metric)["Revenue"].sum() / MainExcelFile["Revenue"].sum()) * 100
  RevenuePercentage.name = "Revenue Percentage"

  Performance = pd.concat([AvgPrice,
                            Quantity,
                            Revenue, 
                            Count,
                            RevenuePercentage], axis=1)

  Performance.loc["Total"] = [MainExcelFile["Unit_Price"].mean().sum(),
                                    MainExcelFile["Quantity"].sum(),
                                      MainExcelFile["Revenue"].sum(),
                                      MainExcelFile["Product"].size,
                                      100,
                                      ] 

  TopQuantity = Quantity.sort_values(ascending=False).head(1)
  TopRevenuePercentage = RevenuePercentage.sort_values(ascending=False).head(1)
  TopRevenue = Revenue.sort_values(ascending=False).head(1)
  TopAvgPrice = AvgPrice.sort_values(ascending=False).head(1)
  TopCount = Count.sort_values(ascending=False).head(1)

  TopPerformance = pd.DataFrame({"Metrics": ["Top Average", "Top Quantity", "Top Revenue", "Top Revenue Percentage", "Most Orders"],
                            "Product":[TopAvgPrice.index[0],TopQuantity.index[0],TopRevenue.index[0],TopRevenuePercentage.index[0],TopCount.index[0]],
                             "Values":[TopAvgPrice.iloc[0],TopQuantity.iloc[0],TopRevenue.iloc[0],TopRevenuePercentage.iloc[0],TopCount.iloc[0]] })

  Performance.to_excel(writer, sheet_name=f"{Metric} Performance")
  TopPerformance.to_excel(writer, sheet_name=f"{Metric} Performance", startrow= len(Performance) + 3, index = False)
     
def SalesSummary(MainExcelFile,writer):
  TotalRevenue = MainExcelFile["Revenue"].sum()
  TotalUnits = MainExcelFile["Quantity"].sum()
  TotalOrders= MainExcelFile.index.size
  AverageOrderValue = MainExcelFile["Revenue"].mean()
  AverageUnitPrice = MainExcelFile["Unit_Price"].mean()
  Summary = pd.DataFrame({"Metric":["TotalRevenue", "TotalUnits", "TotalOrders", "AverageOrderValue","AverageUnitPrice"],
                        "Values":[TotalRevenue,TotalUnits, TotalOrders, AverageOrderValue,AverageUnitPrice ]})
  Summary.to_excel(writer, sheet_name= "Summary")


def main():
  MainExcelFile=DataCleaning()
  with pd.ExcelWriter("Sales Report.xlsx", engine="openpyxl") as writer:
    SalesSummary(MainExcelFile,writer)
    Analysis(MainExcelFile,"Product", writer)
    Analysis(MainExcelFile,"City", writer)
    Analysis(MainExcelFile,"Salesperson", writer)
    Analysis(MainExcelFile,"Month", writer)

if __name__ == '__main__':
  main()