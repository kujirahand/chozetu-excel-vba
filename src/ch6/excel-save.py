# ライブラリを取り込む --- (*1)
import openpyxl as excel

# 新規ワークブックを作る --- (*2)
book = excel.Workbook()

# アクティブなワークシートを得る --- (*3)
sheet = book.active

# A1のセルに値を設定 --- (*4)
sheet["A1"] = "知恵は武器よりも価値がある"

# ファイルを保存 --- (*5)
book.save("hello-py.xlsx")
