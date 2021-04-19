# Excelを起動 --- (*1)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# ブックを開いてシートを得る --- (*2)
$book = $excel.Workbooks.Open($PSScriptRoot + "\hello-ps.xlsx")
$sheet = $book.Sheets(1)

# シートの値を読み込んで表示 --- (*3)
$val = $sheet.Range("A1").Value()
Echo $val

# Excelを終了 --- (*4)
$excel.Quit()

