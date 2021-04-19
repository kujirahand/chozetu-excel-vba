# Excelを起動 --- (*1)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# ブックを追加してシートを取得 --- (*2)
$book = $excel.Workbooks.Add()
$sheet = $book.Sheets(1)

# シートに値を書き込む --- (*3)
$sheet.Range("A1").Value = "こんにちは"

# ファイルに保存 --- (*4)
$fname = $PSScriptRoot + "\hello-ps.xlsx"
$book.SaveAs($fname)

# Excelを終了 --- (*5)
$excel.Quit()

