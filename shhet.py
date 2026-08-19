import openpyxl
# from pyspark.sql import SparkSession
import pandas as pd



# excel_path = r"C:\Users\HP\Desktop\match_report.xlsx"

# xl = pd.ExcelFile(excel_path)
# print(xl.sheet_names)  # Output: ['Sheet1', 'Sheet2', 'Sheet3']

import pandas as pd

excel_path = r"C:\Users\HP\Desktop\match_report.xlsx"

xl = pd.read_excel(excel_path,sheet_name=None)
for a ,b in xl.items():
    print(a)


# for sheet  in xl.sheet_names:
#     df = pd.read_excel(excel_path,sheet_name=sheet)
#     print(df.shape)
#     print(f"hello {sheet}")
#     df.to_csv(f"C:\\Users\\HP\\Desktop\\zero1\\{sheet}.csv" )
#     print(df.head())