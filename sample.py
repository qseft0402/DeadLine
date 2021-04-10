import re
import pandas as pd
df = pd.read_excel("D:\PythonProjects\DeadLine\\test.xlsx")
# excel = pd.ExcelFile("test.xlsx")
# print(excel.sheet_names)
colName = df.columns
print(type(colName))
print(colName)
print(list(colName[2:len(colName)]))

print(colName[2:len(colName)])
# a=[1,2,3]
# b=['a',b'']
# for index, row in df.iterrows():
#
#     print(index)
#     print(tempList)
