# import openpyxl
# import numpy as np
# import datetime
# from openpyxl import load_workbook
# from datetime import datetime
#
# # Replace 'path/to/your/file.xlsx' with the actual path to your XLSX file
# file_path = '/Users/hadi/PycharmProjects/JManTech500/slides_and_excel/230619 UpReach Bootcamp PowerBI and Python Challenge.xlsx'
#
# # Load the Excel workbook
# workbook = openpyxl.load_workbook(file_path)
#
# # Select the first sheet (you can select other sheets by name if needed)
# sheetCust = workbook["Customers"]
# CustomerDict={}
# counter=0
# # Iterate through the rows and columns and read the data
# for row in sheetCust.iter_rows(min_row=1, values_only=True):
#     if row[1] in CustomerDict:
#         try:
#             CustomerDict[row[1]].append((round(row[0])))
#         except:
#             print(" ")
#     else:
#         if counter !=0:
#             CustomerDict[row[1]]=[round(row[0])]
#         else:
#             print(" ")
#         counter+=1
# SalesDict={}
# sheetSales = workbook["Sales"]
#
# for row in sheetSales.iter_rows(min_row=1, values_only=True):
#     try:
#         orderID=row[2]
#         year=str((row[4]).year)
#         SalesDict[orderID] = year
#     except:
#         print("error in type from row 1")
# # Close the workbook after reading
#
# bothTrue=0
#
# only2020=0
# workbook.close()
# TotalCustomers=(len(CustomerDict))
# for i in CustomerDict:
#     Bool2019 = False
#     Bool2020 = False
#     in2020 = False
#     OrdersFromCustomer=(len(CustomerDict[i]))
#
#     if OrdersFromCustomer>=1:
#         for x in (CustomerDict[i]):
#             if x in SalesDict:
#                 year=(SalesDict[x])
#             if year=='2019':
#                 Bool2019=True
#             elif year=='2020':
#                 Bool2020=True
#     if Bool2019==True and Bool2020==True:
#         bothTrue+=1
#     elif Bool2019==False and Bool2020==True:
#         only2020+=1
# print((bothTrue/TotalCustomers)*100)
# print((only2020/TotalCustomers)*100)

## Intermidate
# Date1in=input("enter the start date in formate YYYY-MM-DD")
# Date1=datetime.strptime(Date1in, '%Y-%m-%d')
# Date2in=input("enter the end date in formate YYYY-MM-DD")
# Date2=datetime.strptime(Date1in, '%Y-%m-%d')
#
#
#
# # Replace 'path/to/your/file.xlsx' with the actual path to your XLSX file
# file_path = '/Users/hadi/PycharmProjects/JManTech500/slides_and_excel/230619 UpReach Bootcamp PowerBI and Python Challenge.xlsx'
#
# # Load the Excel workbook
# workbook = openpyxl.load_workbook(file_path)
#
# # Select the first sheet (you can select other sheets by name if needed)
# sheetCust = workbook["Customers"]
# CustomerDict={}
# counter=0
# # Iterate through the rows and columns and read the data
# for row in sheetCust.iter_rows(min_row=1, values_only=True):
#     if row[1] in CustomerDict:
#         try:
#             CustomerDict[row[1]].append((round(row[0])))
#         except:
#             print(" ")
#     else:
#         if counter !=0:
#             CustomerDict[row[1]]=[round(row[0])]
#         else:
#             print(" ")
#         counter+=1
# SalesDict={}
# sheetSales = workbook["Sales"]
#
# for row in sheetSales.iter_rows(min_row=1, values_only=True):
#     try:
#         orderID=row[2]
#         Date=(row[4])
#         SalesDict[orderID] = Date
#     except:
#         print("error in type from row 1")
# # Close the workbook after reading
#
# Inbetween=0
#
# only2020=0
# workbook.close()
# TotalCustomers=(len(CustomerDict))
# for i in CustomerDict:
#     Within = False
#     in2020 = False
#     OrdersFromCustomer=(len(CustomerDict[i]))
#
#     if OrdersFromCustomer>=1:
#         for x in (CustomerDict[i]):
#             if x in SalesDict:
#                 Date=(SalesDict[x])
#             if Date>=Date1 and Date<=Date2:
#                 Within=True
#     if Within==True:
#         Inbetween+=1
# print((Inbetween/TotalCustomers)*100)