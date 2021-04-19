# Excelを起動 --- (*1)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# ブックを追加してシートを取得 --- (*2)
$fname = $PSScriptRoot + "\macro.xlsm"
$book = $excel.Workbooks.Open($fname)

# マクロを実行 --- (*3)
$excel.Application.Run("Sheet1.二倍表示",50)
$excel.Quit()
