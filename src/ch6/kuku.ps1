# Excelを起動 --- (*1)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# ブックを追加してシートを取得 --- (*2)
$book = $excel.Workbooks.Add()
$sheet = $book.Sheets(1)

# シートに九九を書き込む --- (*3)
for ($i = 1; $i -le 9; $i++) {
    for ($j = 1; $j -le 9; $j++) {
        $sheet.Cells($i, $j).Value = $i * $j       
    }
}

# ファイルに保存 --- (*4)
$book.SaveAs($PSScriptRoot + "\kuku.xlsx")
$excel.Quit()

# メモリを開放 --- (*5)
$excel, $book, $sheet | foreach { $_ = $null }

