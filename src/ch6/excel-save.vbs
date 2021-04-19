Dim excel, book, sheet, fso, path
' Excelを起動 --- (*1)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' 新規ブックのシートを得る --- (*2)
Set book = excel.Workbooks.Add()
Set sheet = book.Sheets(1)
' シートのA1に値を代入 --- (*3)
sheet.Range("A1").Value = "こんにちは"
' スクリプトのフォルダにブックを保存 --- (*4)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)
book.SaveAs(path & "\hello.xlsx")
' Excelを閉じる --- (*5)
excel.Quit



