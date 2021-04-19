Dim excel, book, sheet, fso, path
' Excelを起動 --- (*1)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' ブックを開く --- (*2)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(path & "\hello.xlsx")
' シートを得る --- (*3)
Set sheet = book.Sheets(1)
' シートのA1の値を取得 --- (*4)
MsgBox sheet.Range("A1").Value
' Excelを閉じる --- (*5)
excel.Quit



