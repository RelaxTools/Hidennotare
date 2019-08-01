Attribute VB_Name = "Sample"
Option Private Module
Option Explicit

Sub ICursor_Sample()

    'ICursor インターフェースを使用する
    Dim IC As ICursor

    '3行目から↓に向かって読む。
    'B列が空文字列("")になったら終了。
    Set IC = Constructor(New SheetCursor, Sheet1, 3, "B")

    Do Until IC.Eof

        '引数は列を表す文字か列番号を指定する。
        'IC.Item("C").Value でも IC.Item(3).Value でも良い。Rangeを返却。
        Debug.Print IC("C")
        Debug.Print IC("D")
        Debug.Print IC("E")
        
        IC.MoveNext
    Loop

End Sub

Sub StringBuilder_Sample()

    Dim SB As StringBuilder
    
    Set SB = New StringBuilder
    
    SB.Append "A"
    SB.Append "B"
    SB.Append "C"
    SB.Append "D"
    SB.Append "E"

    '文字列の連結
    Debug.Print SB.ToString

    
    Dim S2 As StringBuilder
    
    Set S2 = New StringBuilder
    
    'Trueをつけるとダブルコーテーションで囲む
    S2.Append "red", True
    S2.Append "green", True
    S2.Append "blue", True

    '文字列の連結（カンマ区切り）
    Debug.Print S2.ToString(",", "[", "]")


End Sub

Sub Message_Sample()


    'Information メッセージ
    message.Information "サンプルです。"

    '改行する場合
    message.Information "サンプルです。\n改行も簡単に使えます。"

    'リプレースホルダを使用する場合
    message.Information "サンプルです。{0}のだけでなく{1}もある", "金", "名誉"
    
    'ステータスバー
    message.StatusBar "サンプルです。{0}のだけでなく{1}もある", "金", "名誉"

    
    'リプレースホルダの文字列を返却
    Debug.Print message.PlaceHolder("サンプルです。{0}のだけでなく{1}もある", "金", "名誉")


End Sub

Sub Using_Sample()

    'IUsing に対応したクラスを Usingクラスのコンストラクタに
    '指定することにより、開始・終了をマネジメントする。
    
    'NewExcel           ・・別プロセスでExcel起動・終了を行う。
    'OneTimeSpeedBooster・・再計算、ScreenUpdating及びPrintCommunicationなどを
    '                       停止・再開を行う。
    
    'Withで開始処理、End Withで終了処理を行う。C#でのUsing句のような動作を行う。
    With Constructor(New Using, New NewExcel, New OneTimeSpeedBooster)
    
        'この間で処理を行う。
        Debug.Print "Application.ScreenUpdating:" & Application.ScreenUpdating
    
        'Using クラスの引数１つ目のインスタンスを返す。
        Debug.Print .Args(1).GetInstance.Caption
        

    End With
    '終了
    
    Debug.Print "Application.ScreenUpdating:" & Application.ScreenUpdating

End Sub
Sub Serialize_Sample()

    Dim Row As IList
    Dim Col As IDictionary
    
    Set Row = New ArrayList
    
    Set Col = New Dictionary
    
    Col.Add "Field01", 10
    Col.Add "Field02", 20
    Col.Add "Field03", 30

    Row.Add Col

    Set Col = New Dictionary
    Col.Add "Field01", 40
    Col.Add "Field02", 50
    Col.Add "Field03", 60

    Row.Add Col
    
    Debug.Print Row.ToString

End Sub
Sub desirialize_Sample()

    Dim Row As IList

    Set Row = JSON.ParseJSON("[{""Field01"":10, ""Field02"":20, ""Field03"":30}, {""Field01"":40, ""Field02"":50, ""Field03"":60}]")

    Debug.Print Row.ToString

End Sub

Sub SortedDictionary_sample()


    Dim d As IDictionary
    Dim v As Variant
'    Set D = New SortedDictionary
    Set d = Constructor(New SortedDictionary, New ExplorerComparer)
    
    d.Add "10", "10"
    d.Add "1", "1"
    d.Add "2", "2"

    For Each v In d.Keys
        Debug.Print v
    Next


End Sub
Sub OrderedDictionary_sample()


    Dim d As IDictionary
    Dim v As Variant
    Set d = New OrderedDictionary
    
    d.Add "10", "10"
    d.Add "1", "1"
    d.Add "2", "2"

    d.Remove "1"
    
    d.Key("2") = "3"

    For Each v In d.Keys
        Debug.Print v
    Next
    
'    D.Remove "2"
    d.Remove "10"


End Sub
Sub CsvParser_Sample()

    Dim strBuf As String
    Dim Row As Collection
    Dim Col As Collection
    Dim v As Variant
    strBuf = "1, Watanabe, Fukushima, 36, ""カンマがあっても,OK""" & vbCrLf & "2, satoh, chiba, 24, ""改行があっても" & vbLf & "OKやで"""

    Set Row = StringHelper.CsvParser(strBuf, True)

    For Each Col In Row
        For Each v In Col
            Debug.Print v
        Next
    Next

End Sub
Sub ArrayList_ParseFromListObject_Sample()

    Dim lst As IList
    Dim dic As IDictionary
    Dim Key As Variant

    Set lst = ArrayList.ParseFromListObject(ActiveSheet.ListObjects(1))

    For Each dic In lst

        For Each Key In dic.Keys
        
            Debug.Print dic.Item(Key)
        
        Next

    Next

    Dim a As ArrayList
    
    Set a = lst
    
    a.CopyToListObject ActiveSheet.ListObjects(2)

End Sub
Sub Dictionary_ParseFromListObject_Sample()

    Dim lst As IDictionary
    Dim dic As IDictionary
    Dim Key As Variant
    Dim Key2 As Variant

    Set lst = Dictionary.ParseFromListObject(ActiveSheet.ListObjects(1), "A")

    For Each Key In lst.Keys

        Set dic = lst.Item(Key)

        For Each Key2 In dic.Keys
        
            Debug.Print dic.Item(Key2)
        
        Next

    Next

    Dim a As Dictionary
    
    Set a = lst
    
    a.CopyToListObject ActiveSheet.ListObjects(2)

End Sub

Sub IsDictionary_Sample()

    Dim dic As Object
    
    
    Set dic = New Dictionary
    
    Debug.Print IsDictionary(dic)

    Set dic = New OrderedDictionary
    
    Debug.Print IsDictionary(dic)

    Set dic = New SortedDictionary
    
    Debug.Print IsDictionary(dic)

    Set dic = CreateObject("Scripting.Dictionary")
    
    Debug.Print IsDictionary(dic)


End Sub
Sub Web()

    'http://weather.livedoor.com/weather_hacks/webservice
    Dim strBuf As String
    Dim v As IDictionary
    
    strBuf = Application.WorksheetFunction.WebService("http://weather.livedoor.com/forecast/webservice/json/v1?city=120010")

    
    Dim dic As IDictionary
    
    Set dic = JSON.ParseJSON(strBuf)
'    Debug.Print strBuf

    Dim lst As IList
    Set lst = dic.Item("forecasts")

    For Each v In lst
    
        Debug.Print v.Item("date")
        Debug.Print v.Item("dateLabel")
        Debug.Print v.Item("telop")
        If IsEmpty(v.Item("temperature").Item("max")) Then
            Debug.Print ""
        Else
           Debug.Print v.Item("temperature").Item("max").Item("celsius")
        End If
        If IsEmpty(v.Item("temperature").Item("min")) Then
            Debug.Print ""
        Else
           Debug.Print v.Item("temperature").Item("min").Item("celsius")
        End If
    
    Next


End Sub
