Attribute VB_Name = "mecab"
Option Explicit

'
' MeCab for Excel VBA (Windows)
'

Dim MeCabPath As String
Dim MeCabCharset As String
Dim MeCabOptions As String
Dim MeCabDictDir As String
Const MeCabCharsetDefault = "Shift_JIS"

' MeCabの結果データを保持するユーザー型
Public Type MeCabItem
    表層形 As String
    品詞 As String
    品詞詳細1 As String
    品詞詳細2 As String
    品詞詳細3 As String
    活用形 As String
    活用型 As String
    原形 As String
    ヨミ As String
    発音 As String
End Type


Public Sub SetMeCabPath(ByVal Path)
    MeCabPath = Path
End Sub

Public Sub SetMeCabCharset(ByVal Charset As String)
    MeCabCharset = Charset
End Sub

Public Sub SetMeCabOptions(ByVal Options As String)
    MeCabOptions = Options
End Sub

Public Sub SetMeCabDictDir(ByVal DictDir As String)
    MeCabDictDir = DictDir
End Sub

Private Sub MeCabInit()
    ' Find MeCab
    If MeCabPath = "" Then
        MeCabPath = "C:\Program Files (x86)\MeCab\bin\mecab.exe"
        If FileExists(MeCabPath) = False Then
            MeCabPath = "C:\Program Files\MeCab\bin\mecab.exe"
            If FileExists(MeCabPath) = False Then
                MeCabPath = ThisWorkbook.Path & "\mecab.exe"
                If FileExists(MeCabPath) = False Then
                    MeCabPath = ThisWorkbook.Path & "\bin\mecab.exe"
                    If FileExists(MeCabPath) = False Then
                        MsgBox "MeCabがインストールされていません。" & vbCrLf & _
                            "あるいは、MeCabのパスを指定してください。"
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub MeCabExecToSheet(ByVal InText As String, ByRef Sheet As Worksheet, ByVal Top As Integer)
    Dim Res As String
    Res = MeCabExec(InText)
    If Res = "" Then Exit Sub
    Dim Lines, y, x
    Lines = Split(Res, vbCrLf)
    For y = 0 To UBound(Lines)
        Dim Line As String
        Line = Lines(y)
        
        Dim Tabs
        Tabs = Split(Line, Chr(9))
        If UBound(Tabs) = 0 Then Exit For
        Dim Word, Desc
        Word = Tabs(0)
        Desc = Tabs(1)
        
        Dim Cm
        Cm = Split(Desc, ",")
        If UBound(Cm) < 8 Then Exit For
        
        Dim row
        row = y + Top
        
        Sheet.Cells(row, 1) = Word
        For x = 0 To UBound(Cm)
            Sheet.Cells(row, 2 + x) = Cm(x)
        Next
    Next
End Sub


Public Function MeCabExecToItems(ByVal InText As String) As MeCabItem()
    Dim Res As String
    Res = Trim(MeCabExec(InText))
    Dim a, i
    a = Split(Res, vbCrLf)
    Dim items() As MeCabItem
    ReDim items(UBound(a) + 1)
    For i = 0 To UBound(a)
        If a(i) = "" Then Exit For
        Dim rowa, Word, Desc, da
        ' tab
        rowa = Split(a(0), Chr(9))
        If UBound(rowa) = 0 Then Exit For
        Word = rowa(0)
        Desc = rowa(1)
        ' comma : 品詞,品詞細分類1,品詞細分類2,品詞細分類3,活用型,活用形,原形,読み,発音
        da = Split(Desc, ",")
        If UBound(da) < 8 Then Exit For
        items(i).表層形 = Word
        items(i).品詞 = da(0)
        items(i).品詞詳細1 = da(1)
        items(i).品詞詳細2 = da(2)
        items(i).品詞詳細3 = da(3)
        items(i).活用型 = da(4)
        items(i).活用形 = da(5)
        items(i).原形 = da(6)
        items(i).ヨミ = da(7)
        items(i).発音 = da(8)
    Next
    MeCabExecToItems = items
End Function


Public Function MeCabExec(ByVal InText As String) As String
    Dim InFile As String, ResultFile As String, Cmd As String, Res As String
    Dim BatFile As String, Opt As String
    
    ' MeCabの初期化
    Call MeCabInit
    
    BatFile = GetTempPath(".bat")
    InFile = GetTempPath(".txt")
    ResultFile = GetTempPath(".txt")
    
    ' 入力テキストをファイルに保存
    MeCabSaveText InFile, InText ' 入力テキストはインストール辞書のコード
    
    ' オプションを反映
    Opt = "" & MeCabOptions
    If MeCabDictDir <> "" Then
        Opt = Opt & " -d " & MeCabDictDir
    End If
    
    ' バッチを作成
    Cmd = "type " & qq(InFile) & " | " & qq(MeCabPath) & " " & MeCabOptions & " > " & qq(ResultFile) & vbCrLf
    ' Cmd = Cmd & "pause" & vbCrLf
    SaveToFile BatFile, Cmd, "Shift_JIS" ' バッチファイルはShift_JIS必須
    Debug.Print Cmd
    
    ' バッチを実行
    If ShellWait(BatFile) Then
        Res = MeCabLoadText(ResultFile)
        ' Debug.Print Res
        MeCabExec = Res
    Else
        MeCabExec = ""
    End If
End Function


Private Function GetTempPath(Ext As String) As String
    Dim FSO As Object, tmp As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tmp = FSO.GetSpecialFolder(2) & "\" & FSO.GetBaseName(FSO.GetTempName) & Ext
    GetTempPath = tmp
End Function


' Clear sheet
Public Sub ClearSheet(ByRef Sheet As Worksheet, ByVal TopRow As Integer)
    Dim EndCol, EndRow, row, Col
    With Sheet.UsedRange
        EndRow = .Rows(.Rows.Count).row
        EndCol = .Columns(.Columns.Count).Column
    End With
    For row = TopRow To EndRow
        For Col = 1 To EndCol
            Sheet.Cells(row, Col) = ""
        Next
    Next
End Sub

' TSV to Sheet
Public Sub TSVToSheet(ByRef Sheet As Worksheet, ByVal tsv As String, TopRow As Integer)
    Dim Rows As Variant, Cols As Variant
    Dim i, j
    Rows = Split(tsv, Chr(10))
    For i = 0 To UBound(Rows)
        Cols = Split(Rows(i), Chr(9))
        For j = 0 To UBound(Cols)
            Dim v
            v = Cols(j)
            v = Replace(v, "¶", vbCrLf)
            Sheet.Cells(i + TopRow, j + 1) = v
        Next
    Next
End Sub


' Convert Sheet to TSV
Public Function ToTSV(ByRef Sheet As Worksheet) As String
    Dim s As String
    s = ""
    ' シートの範囲を取得
    Dim BottomRow As Integer, RightCol As Integer
    BottomRow = Sheet.Range("A1").End(xlDown).row
    RightCol = Sheet.Range("A1").End(xlToRight).Column
    ' シート範囲を左上から順に取得
    Dim y, x, v
    For y = 1 To BottomRow
        For x = 1 To RightCol
            v = Sheet.Cells(y, x)
            ' セル内の改行だけは置換しておく
            v = Replace(v, vbCrLf, "¶")
            s = s & v & Chr(9)
        Next
        s = s & vbCrLf
    Next
    ToTSV = s
End Function

' クォート処理
Public Function qq(str) As String
    qq = """" & str & """"
End Function

' コマンドを実行して終了まで待機する
Public Function ShellWait(ByVal Command As String) As Boolean
    On Error GoTo SHELL_ERROR
    Dim wsh As Object
    Set wsh = CreateObject("Wscript.Shell")
    Dim Res As Integer
    Res = wsh.Run(Command, 7, True) ' 最小化して終了まで待機
    ShellWait = (Res = 0)
    Exit Function
SHELL_ERROR:
    ShellWait = False
End Function

Public Sub MeCabSaveText(ByVal Filename, ByVal Text)
    If MeCabCharset = "" Then MeCabCharset = MeCabCharsetDefault
    SaveToFile Filename, Text, MeCabCharset
End Sub

Public Function MeCabLoadText(ByVal Filename) As String
    If MeCabCharset = "" Then MeCabCharset = MeCabCharsetDefault
    MeCabLoadText = LoadFromFile(Filename, MeCabCharset)
End Function

' 任意の文字エンコーディングを指定してテキストファイルを読む
Private Function LoadFromFile(Filename, Charset) As String
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = Charset
    stream.Open
    stream.LoadFromFile Filename
    LoadFromFile = stream.ReadText
    stream.Close
End Function

' テキストを指定文字コードでファイルに保存
Private Sub SaveToFile(ByVal Filename, ByVal Text, ByVal Charset)
    ' UTF-8 の場合 BOMは不要
    If LCase(Charset) = "utf-8" Or LCase(Charset) = "utf-8n" Or LCase(Charset) = "utf8" Then
        Call SaveToFileUTF8N(Filename, Text)
        Exit Sub
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = Charset
    stream.Open
    stream.WriteText Text
    stream.SaveToFile Filename, 2
    stream.Close
End Sub

' BOMなしのUTF-8でファイルにテキストを書き込む
Private Sub SaveToFileUTF8N(Filename, Text)
    Dim stream, buf
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' テキストモードを指定 --- (*1)
        .Charset = "UTF-8"
        .Open
        .WriteText Text ' テキストを書き込む
        .Position = 0 ' カーソルをファイル先頭に --- (*2)
        .Type = 1 ' バイナリモードに変更
        .Position = 3 ' BOM(3バイト)を飛ばす
        buf = .Read() ' 内容を読み込む
        .Position = 0 ' カーソルを先頭に --- (*3)
        .Write buf ' BOMなしのテキストを書き込み
        .SetEOS
        .SaveToFile Filename, 2
        .Close
    End With
End Sub

Private Function FileExists(ByVal Filename) As Boolean
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileExists = FSO.FileExists(Filename)
    Set FSO = Nothing
End Function

