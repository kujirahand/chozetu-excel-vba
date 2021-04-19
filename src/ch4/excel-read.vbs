' Excelを起動
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' 既存ブックを読む
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(ThisPath & "\test.xlsx")
' 操作するシートを選択
Set sheet = book.Worksheets("Sheet1")
' セルA1の値を読み取る
v = sheet.Range("A1").Value
MsgBox v
book.Close
excel.Quit
