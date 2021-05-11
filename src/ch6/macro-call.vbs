' Excelを起動
Set excel = CreateObject("Excel.Application")
excel.Visible = True
' マクロを定義したブックを読む
set fso = createObject("Scripting.FileSystemObject")
ThisPath = fso.getParentFolderName(WScript.ScriptFullName)
Set book = excel.Workbooks.Open(ThisPath & "\macro-message-call.xlsm")
' マクロを起動
excel.Application.Run "Sheet1.ShowMessage"
' Excelを終了
excel.Quit
Set excel = Nothing
