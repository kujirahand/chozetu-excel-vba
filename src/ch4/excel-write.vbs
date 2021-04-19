' Excelを起動
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' 新規ブックを作成
Set book = excel.Workbooks.Add()
' 操作するシートを選択
Set sheet = book.Worksheets("Sheet1")
' セルA1に格言を値を入れる
sheet.Range("A1").Value = "怠け者は何も得ず勤勉な人は満たされる"
' 名前を付けて保存
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
book.SaveAs ThisPath & "\test.xlsx"
' Excelを閉じる
excel.Quit
