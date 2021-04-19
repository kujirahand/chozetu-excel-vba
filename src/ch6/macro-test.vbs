Dim excel, book, fso, path
' VBScriptのあるパスを調べる --- (*1)
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(WScript.ScriptFullName)

' Excelを起動してブックを開く --- (*2)
Set excel = CreateObject("Excel.Application")
excel.Visible = True
Set book = excel.Workbooks.Open(path & "\macro.xlsm")

' マクロを実行 --- (*3)
excel.Application.Run "Sheet1.二倍表示", 30

' Excelを閉じる --- (*4)
excel.Quit
