Attribute VB_Name = "Test"
Option Explicit
Option Private Module

' 全テストコマンド
Sub Test_All()

    Call Test_ArrayList
    Call Test_LinkedList
    Call Test_ICursor_SheetCursor
    Call Test_CollectionCursor
    Call Test_LineCursor
    Call Test_StringBuilder
    Call Test_Serialize
    Call Test_desirialize
    Call Test_SortedDictionary
    Call Test_OrderedDictionary
    Call Test_CsvParser
    Call Test_IsDictionary
    Call Test_MCommand
    Call Test_TextWriter
    Call Test_CsvWriter
    Call Test_Compress
    Call Test_PlaceHolder

    Message.Information "正常終了!"

End Sub


Sub Test_ArrayList()

    Dim IL As IList
    Dim IC As ICursor
    
    Set IL = New ArrayList
    
    Debug.Assert IL.Count = 0
    
    IL.Add "a"
    IL.Add 1
    IL.Add 3.14159

    Debug.Assert IL.Item(0) = "a"
    Debug.Assert IL.Item(1) = 1
    Debug.Assert IL.Item(2) = 3.14159

    Debug.Assert IL.Count = 3

    IL.RemoveAt 1
    
    Debug.Assert IL.Item(0) = "a"
    Debug.Assert IL.Item(1) = 3.14159
    
    Debug.Assert IL.Count = 2
        
    IL.Clear
    
    Debug.Assert IL.Count = 0
    
    Dim col As Collection
    
    Set col = New Collection
    
    col.Add "a"
    col.Add 1
    col.Add 3.14159
    
    Set IC = ArrayList.CreateObject(col)
    Dim i As Long
    
    i = 0
    Do Until IC.Eof
    
        Select Case i
            Case 0
                Debug.Assert IC.Item = "a"
            Case 1
                Debug.Assert IC.Item = 1
            Case 2
                Debug.Assert IC.Item = 3.14159
        End Select
        
        i = i + 1
        IC.MoveNext
    
    Loop
    
    Debug.Assert i = 3
    
    Set IL = IC
    
    Debug.Assert IL.Count = 3
    
    IL.Clear
    
    Debug.Assert IL.Count = 0
    
    Dim v As Variant
    
    v = Array("a", 1, 3.14159)
    
    Set IL = ArrayList.CreateObject(v)
    
    'JSON
    Debug.Assert IL.ToString = "[""a"", 1, 3.14159]"
    
    Debug.Assert Join(IL.ToArray(), "/") = "a/1/3.14159"

    IL.Sort

    Debug.Assert IL.Item(0) = 1
    Debug.Assert IL.Item(1) = 3.14159
    Debug.Assert IL.Item(2) = "a"
    
    Debug.Assert IL.Count = 3
    
    IL.Insert 1, "追加"

    Debug.Assert IL.Item(0) = 1
    Debug.Assert IL.Item(1) = "追加"
    Debug.Assert IL.Item(2) = 3.14159
    Debug.Assert IL.Item(3) = "a"
    
    Debug.Assert IL.Count = 4
    


End Sub
Sub Test_LinkedList()

    Dim IL As IList
    Dim IC As ICursor
    
    Set IL = New LinkedList
    
    Debug.Assert IL.Count = 0
    
    IL.Add "a"
    IL.Add 1
    IL.Add 3.14159

    Debug.Assert IL.Item(0) = "a"
    Debug.Assert IL.Item(1) = 1
    Debug.Assert IL.Item(2) = 3.14159

    Debug.Assert IL.Count = 3

    IL.RemoveAt 1
    
    Debug.Assert IL.Item(0) = "a"
    Debug.Assert IL.Item(1) = 3.14159
    
    Debug.Assert IL.Count = 2
        
    IL.Clear
    
    Debug.Assert IL.Count = 0
    
    Dim col As Collection
    
    Set col = New Collection
    
    col.Add "a"
    col.Add 1
    col.Add 3.14159
    
    Set IC = LinkedList.CreateObject(col)
    Dim i As Long
    
    i = 0
    Do Until IC.Eof
    
        Select Case i
            Case 0
                Debug.Assert IC.Item = "a"
            Case 1
                Debug.Assert IC.Item = 1
            Case 2
                Debug.Assert IC.Item = 3.14159
        End Select
        
        i = i + 1
        IC.MoveNext
    
    Loop
    
    Debug.Assert i = 3
    
    Set IL = IC
    
    Debug.Assert IL.Count = 3
    
    IL.Clear
    
    Debug.Assert IL.Count = 0
    
    Dim v As Variant
    
    v = Array("a", 1, 3.14159)
    
    Set IL = LinkedList.CreateObject(v)
    
    'JSON
    Debug.Assert IL.ToString = "[""a"", 1, 3.14159]"
    
    Debug.Assert Join(IL.ToArray(), "/") = "a/1/3.14159"

    IL.Sort

    Debug.Assert IL.Item(0) = 1
    Debug.Assert IL.Item(1) = 3.14159
    Debug.Assert IL.Item(2) = "a"
    
    Debug.Assert IL.Count = 3
    
    IL.Insert 1, "追加"

    Debug.Assert IL.Item(0) = 1
    Debug.Assert IL.Item(1) = "追加"
    Debug.Assert IL.Item(2) = 3.14159
    Debug.Assert IL.Item(3) = "a"
    
    Debug.Assert IL.Count = 4
    


End Sub

Sub Test_ICursor_SheetCursor()

    'ICursor インターフェースを使用する
    Dim IC As ICursor

    '3行目から↓に向かって読む。
    'B列が空文字列("")になったら終了。
    Set IC = SheetCursor.CreateObject(Sheet1, 3, "B")

    Dim i As Long
    i = 0
    Do Until IC.Eof

        '引数は列を表す文字か列番号を指定する。
        'IC.Item("C").Value でも IC.Item(3).Value でも良い。Rangeを返却。
        Select Case i
            Case 0
                Debug.Assert IC.Item("C") = "A1"
                Debug.Assert IC.Item("D") = "B1"
                Debug.Assert IC.Item("E") = "C1"
            Case 1
                Debug.Assert IC.Item("C") = "A2"
                Debug.Assert IC.Item("D") = "B2"
                Debug.Assert IC.Item("E") = "C2"
            Case 2
                Debug.Assert IC.Item("C") = "A3"
                Debug.Assert IC.Item("D") = "B3"
                Debug.Assert IC.Item("E") = "C3"
        End Select
        i = i + 1
        IC.MoveNext
    Loop

End Sub
Sub Test_CollectionCursor()

    Dim col As Collection
    Set col = New Collection
    
    col.Add "a"
    col.Add "b"
    col.Add "c"
    col.Add "D"

    Dim IC As ICursor

    Set IC = CollectionCursor.CreateObject(col)
    Dim i As Long
    i = 0
    Do Until IC.Eof
    
        Select Case i
            Case 0
                Debug.Assert IC.Item = "a"
            Case 1
                Debug.Assert IC.Item = "b"
            Case 2
                Debug.Assert IC.Item = "c"
            Case 3
                Debug.Assert IC.Item = "D"
        End Select
        i = i + 1
        IC.MoveNext
    Loop

End Sub
Sub Test_LineCursor()


    Dim v As Variant
    
    v = Array("a", "b", "c")


    Dim IC As ICursor

    Set IC = LineCursor.CreateObject(v)

    Dim i As Long
    i = 0
    Do Until IC.Eof
    
        Select Case i
            Case 0
                Debug.Assert IC.Item = "a"
            Case 1
                Debug.Assert IC.Item = "b"
            Case 2
                Debug.Assert IC.Item = "c"
        End Select
        i = i + 1
        IC.MoveNext
    Loop

End Sub
Sub Test_StringBuilder()

    Dim SB As StringBuilder
    
    Set SB = New StringBuilder
    
    SB.Append "A"
    SB.Append "B"
    SB.Append "C"
    SB.Append "D"
    SB.Append "E"

    '文字列の連結
    Debug.Assert SB.ToString = "ABCDE"
    
    Dim S2 As StringBuilder
    
    Set S2 = New StringBuilder
    
    'Trueをつけるとダブルコーテーションで囲む
    S2.Append "red", True
    S2.Append "green", True
    S2.Append "blue", True

    '文字列の連結（カンマ区切り）
    Debug.Assert S2.ToString(",", "[", "]") = "[""red"",""green"",""blue""]"

End Sub

Sub Test_Serialize()

    Dim Row As IList
    Dim col As IDictionary
    
    Set Row = New ArrayList
    
    Set col = New Dictionary
    
    col.Add "Field01", 10
    col.Add "Field02", 20
    col.Add "Field03", 30

    Row.Add col

    Set col = New Dictionary
    col.Add "Field01", 40
    col.Add "Field02", 50
    col.Add "Field03", 60

    Row.Add col
    
    Debug.Assert Row.ToString = "[{""Field01"":10, ""Field02"":20, ""Field03"":30}, {""Field01"":40, ""Field02"":50, ""Field03"":60}]"

End Sub
Sub Test_desirialize()

    Dim Row As IList

    Set Row = Parser.ParseJSON("[{""Field01"":10, ""Field02"":20, ""Field03"":30}, {""Field01"":40, ""Field02"":50, ""Field03"":60}]")

    Debug.Assert Row.Count = 2

    Debug.Assert Row.ToString = "[{""Field01"":10, ""Field02"":20, ""Field03"":30}, {""Field01"":40, ""Field02"":50, ""Field03"":60}]"

End Sub

Sub Test_SortedDictionary()


    Dim d As IDictionary
    Dim v As Variant
    
    Set d = SortedDictionary.CreateObject()
    
    d.Add "10", "10"
    d.Add "1", "1"
    d.Add "2", "2"

    Debug.Assert d.Keys(0) = "1"
    Debug.Assert d.Keys(1) = "10"
    Debug.Assert d.Keys(2) = "2"

    Set d = SortedDictionary.CreateObject(New ExplorerComparer)
    
    d.Add "10", "10"
    d.Add "1", "1"
    d.Add "2", "2"

    Debug.Assert d.Keys(0) = "1"
    Debug.Assert d.Keys(1) = "2"
    Debug.Assert d.Keys(2) = "10"


End Sub
Sub Test_OrderedDictionary()


    Dim d As IDictionary
    Dim v As Variant
    Set d = New OrderedDictionary
    
    d.Add "10", "10"
    d.Add "1", "1"
    d.Add "2", "2"
    
    Debug.Assert d.Keys(0) = "10"
    Debug.Assert d.Keys(1) = "1"
    Debug.Assert d.Keys(2) = "2"
    
    d.Remove "1"
    
    Debug.Assert d.Keys(0) = "10"
    Debug.Assert d.Keys(1) = "2"
    
    d.Key("2") = "3"

    Debug.Assert d.Keys(0) = "10"
    Debug.Assert d.Keys(1) = "3"
    
    Dim i As Long
    i = 0
    For Each v In d
        Select Case i
            Case 0
                Debug.Assert d.Keys(i) = "10"
            Case 1
                Debug.Assert d.Keys(i) = "3"
        End Select
        i = i + 1
    Next
    
    d.Remove "10"
    
    Debug.Assert d.Keys(0) = "3"

    Debug.Assert d.Count = 1

End Sub
Sub Test_CsvParser()

    Dim strBuf As String
    Dim Row As Collection
    Dim col As Collection
    Dim v As Variant
    strBuf = "1,Watanabe,Fukushima,36,""カンマがあっても,OK""" & vbCrLf & "2,satoh,chiba,24,""改行があっても" & vbLf & "OKやで"""

    Set Row = Parser.ParseCsv(strBuf, True)

    Debug.Assert Row(1)(1) = "1"
    Debug.Assert Row(1)(2) = "Watanabe"
    Debug.Assert Row(1)(3) = "Fukushima"
    Debug.Assert Row(1)(4) = "36"
    Debug.Assert Row(1)(5) = "カンマがあっても,OK"

    Debug.Assert Row(2)(1) = "2"
    Debug.Assert Row(2)(2) = "satoh"
    Debug.Assert Row(2)(3) = "chiba"
    Debug.Assert Row(2)(4) = "24"
    Debug.Assert Row(2)(5) = "改行があっても" & vbLf & "OKやで"

End Sub
Sub Test_IsDictionary()

    Dim dic As Object
    
    Set dic = New Dictionary
    
    Debug.Assert Core.IsDictionary(dic)

    Set dic = New OrderedDictionary
    
    Debug.Assert Core.IsDictionary(dic)

    Set dic = New SortedDictionary
    
    Debug.Assert Core.IsDictionary(dic)

    Set dic = VBA.CreateObject("Scripting.Dictionary")
    
    Debug.Assert Core.IsDictionary(dic)

    Debug.Assert Core.IsDictionary("") = False

    Dim lst As New ArrayList

    Debug.Assert Core.IsDictionary(lst) = False


End Sub
'------------------------------------------------
' MCommandをVBAで作成する場合のヘルパークラス
'------------------------------------------------
Sub Test_MCommand()

    '-----------------------------------
    ' MCommandを代入せずに作成する場合
    '-----------------------------------
    Dim t1 As MTable
    Dim t2 As MTable
    Dim t3 As MTable
    
    Set t1 = MCsv.Document(MFile.Contents("C:\Test.csv"), _
                           "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    Set t2 = MTable.Skip(t1, 2)
    Set t3 = MTable.PromoteHeaders(t2, "[PromoteAllScalars=true]")

    Dim m1 As MCommand
    Set m1 = New MCommand
    
    m1.Append t3
    
    Dim strBuf As String
    
    strBuf = "let " & vbCrLf
    strBuf = strBuf & "Source1 = Table.PromoteHeaders(Table.Skip(Csv.Document(File.Contents(""C:\Test.csv""), [Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]), 2), [PromoteAllScalars=true]) " & vbCrLf
    strBuf = strBuf & "in Source1"
    
    Debug.Assert m1.ToString = strBuf

    
    '結果
    'let Source1 = Table.PromoteHeaders(Table.Skip(Csv.Document(File.Contents("C:\Test.csv"),
    '              [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]), 2), [PromoteAllScalars=true]) in Source1

    
    '-----------------------------------
    ' MCommandに代入して作成する場合
    '-----------------------------------
    Dim m2 As MCommand
    Set m2 = New MCommand
    
    m2.Append MCsv.Document(MFile.Contents("C:\Test.csv"), _
                            "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    
    m2.Append MTable.Skip(m2.Table, 2)
    m2.Append MTable.PromoteHeaders(m2.Table, "[PromoteAllScalars=true]")

    strBuf = "let " & vbCrLf
    strBuf = strBuf & "Source1 = Csv.Document(File.Contents(""C:\Test.csv""), [Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]), " & vbCrLf
    strBuf = strBuf & "Source2 = Table.Skip(Source1, 2), " & vbCrLf
    strBuf = strBuf & "Source3 = Table.PromoteHeaders(Source2, [PromoteAllScalars=true]) " & vbCrLf
    strBuf = strBuf & "in Source3"
    
    Debug.Assert m2.ToString = strBuf

    '結果
    'let Source1 = Csv.Document(File.Contents("C:\Test.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    '    Source2 = Table.Skip(Source1, 2),
    '    Source3 = Table.PromoteHeaders(Source2, [PromoteAllScalars=true]) in Source3


    '-----------------------------------
    ' MRecord/MListを用いたサンプル
    '-----------------------------------
    Dim m3 As MCommand
    
    'MRecord(M言語のRecord) は DictionaryをWrapしたもの。使用方法はDictionary同等。
    Dim rec As IDictionary
    Set rec = New MRecord
            
    rec.Add "Column1", """No."""
    rec.Add "Column2", """NAME"""
    rec.Add "Column3", """AGE"""
    rec.Add "Column4", """ADDRESS"""
    rec.Add "Column5", """TEL"""
    
    'MList(M言語のList) は CollectionをWrapしたもの。使用方法はCollectionと同等。
    Dim lst As IList
    Set lst = New MList
    lst.Add rec
    
    Set m3 = New MCommand

    m3.Append MCsv.Document(MFile.Contents("C:\Test.csv"), _
                            "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    m3.Append MTable.Skip(m3.Table, 2)
    m3.Append MTable.InsertRows(m3.Table, 0, lst)
    m3.Append MTable.PromoteHeaders(m3.Table, "[PromoteAllScalars=true]")

    strBuf = "let " & vbCrLf
    strBuf = strBuf & "Source1 = Csv.Document(File.Contents(""C:\Test.csv""), [Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]), " & vbCrLf
    strBuf = strBuf & "Source2 = Table.Skip(Source1, 2), " & vbCrLf
    strBuf = strBuf & "Source3 = Table.InsertRows(Source2, 0, {Column1, Column2, Column3, Column4, Column5}), " & vbCrLf
    strBuf = strBuf & "Source4 = Table.PromoteHeaders(Source3, [PromoteAllScalars=true]) " & vbCrLf
    strBuf = strBuf & "in Source4"


    Debug.Assert m3.ToString = strBuf

    '結果
    'let Source1 = Csv.Document(File.Contents("C:\Test.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    '    Source2 = Table.Skip(Source1, 2),
    '    Source3 = Table.InsertRows(Source2, 0, {[Column1="No.", Column2="NAME", Column3="AGE", Column4="ADDRESS", Column5="TEL"]}),
    '    Source4 = Table.PromoteHeaders(Source3, [PromoteAllScalars=true]) in Source4

End Sub

Sub Test_TextWriter()

    Dim strFile As String
    Dim strBuf As String
    
    strFile = FileIO.BuildPath(ThisWorkbook.Path, "testxx.txt")

    '空ファイル
    With TextWriter.CreateObject(strFile, NewLineCodeLF, EncodeUTF8, OpenModeOutput, False)
    End With

    Dim blnFind As Boolean
    blnFind = False

    With TextReader.CreateObject(strFile, NewLineCodeLF, EncodeUTF8)

        Do Until .Eof

            blnFind = True

            .MoveNext
        Loop

    End With
    
    Debug.Assert blnFind = False
    
    
    With TextWriter.CreateObject(strFile, NewLineCodeLF, EncodeUTF8, OpenModeOutput, False)

        .WriteLine "あいうえお"

    End With


    With TextReader.CreateObject(strFile, NewLineCodeLF, EncodeUTF8)

        Do Until .Eof

            strBuf = .Item

            .MoveNext
        Loop

    End With
    
    FileIO.DeleteFile strFile

    Debug.Assert strBuf = "あいうえお"

End Sub
Sub Test_CsvWriter()

    Dim strFile As String
    Dim strBuf1 As String
    Dim strBuf2 As String
    
    strFile = FileIO.BuildPath(ThisWorkbook.Path, "testxx.csv")
    
    '空ファイル
    With CsvWriter.CreateObject(strFile, NewLineCodeLF, EncodeUTF16LE, OpenModeOutput, True, ",")
    End With

    Dim blnFind As Boolean
    blnFind = False
    
    With CSVReader.CreateObject(strFile, NewLineCodeLF, EncodeUTF16LE, ",", True)
        Do Until .Eof
            blnFind = True
            .MoveNext
        Loop

    End With

    Debug.Assert blnFind = False


    With CsvWriter.CreateObject(strFile, NewLineCodeLF, EncodeUTF16LE, OpenModeOutput, True, ",")

        .WriteLine Array("あい,うえお", Core.PlaceHolder("かきく\nけこ"))

    End With


    With CSVReader.CreateObject(strFile, NewLineCodeLF, EncodeUTF16LE, ",", True)

        Do Until .Eof

            strBuf1 = .Item(1)
            strBuf2 = .Item(2)

            .MoveNext
        Loop

    End With
    
    FileIO.DeleteFile strFile

    Debug.Assert strBuf1 = "あい,うえお"
    Debug.Assert strBuf2 = "かきく" & vbLf & "けこ"

End Sub
Sub Test_Compress()

    Dim strTmp  As String
    Dim strFile As String
    Dim strZip As String
    
    strTmp = FileIO.TempFolder
    strFile = FileIO.BuildPath(strTmp, "aaa.txt")
    strZip = FileIO.BuildPath(strTmp, "aaa.zip")

    TextWriter.CreateObject(strFile).WriteLine ("ああああ")


    Dim lst As IList
    
    Set lst = New ArrayList
    
    lst.Add strFile

    Call Zip.CompressArchive(lst.ToArray, strZip)

    Debug.Assert FileIO.FileExists(strZip)

    FileIO.DeleteFile strFile
    
    Call Zip.ExpandArchive(strZip, strTmp)

    Debug.Assert FileIO.FileExists(strFile)
    
    FileIO.DeleteFile strFile
    FileIO.DeleteFile strZip

End Sub
Sub Test_PlaceHolder()

    Debug.Assert Core.PlaceHolder("これはテストです。\n{0}", 10) = "これはテストです。" & vbLf & "10"

End Sub
