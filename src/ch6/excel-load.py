import openpyxl as excel

# ワークブックを開く --- (*1)
book = excel.load_workbook("hello-py.xlsx")

# 先頭のワークシートを参照する --- (*2)
sheet = book.worksheets[0]

# A1のセルの値を得る --- (*3)
v = sheet["A1"].value

# 画面に表示 --- (*4)
print(v)
