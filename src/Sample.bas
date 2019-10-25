Attribute VB_Name = "Sample"
Option Private Module
Option Explicit


Sub Message_Sample()


    'Information メッセージ
    Message.Information "サンプルです。"

    '改行する場合
    Message.Information "サンプルです。\n改行も簡単に使えます。"

    'リプレースホルダを使用する場合
    Message.Information "サンプルです。{0}のだけでなく{1}もある", "金", "名誉"
    
    'ステータスバー
    Message.StatusBar "サンプルです。{0}のだけでなく{1}もある", "金", "名誉"

    
    'リプレースホルダの文字列を返却
    Debug.Print Core.PlaceHolder("サンプルです。{0}のだけでなく{1}もある", "金", "名誉")


End Sub

Sub Using_Sample()

    'IUsing に対応したクラスを Usingクラスのコンストラクタに
    '指定することにより、開始・終了をマネジメントする。
    
    'NewExcel           ・・別プロセスでExcel起動・終了を行う。
    'OneTimeSpeedBooster・・再計算、ScreenUpdating及びPrintCommunicationなどを
    '                       停止・再開を行う。
    
    'Withで開始処理、End Withで終了処理を行う。C#でのUsing句のような動作を行う。
    With Using.CreateObject(New NewExcel, New OneTimeSpeedBooster)
    
        'この間で処理を行う。
        Debug.Print "Application.ScreenUpdating:" & Application.ScreenUpdating
    
        'Using クラスの引数１つ目のインスタンスを返す。
        Debug.Print .Args(1).GetInstance.Caption
        

    End With
    '終了
    
    Debug.Print "Application.ScreenUpdating:" & Application.ScreenUpdating

End Sub


Sub Web()

    'http://weather.livedoor.com/weather_hacks/webservice
    Dim strBuf As String
    Dim v As IDictionary
    
    strBuf = Application.WorksheetFunction.WebService("http://weather.livedoor.com/forecast/webservice/json/v1?city=120010")

    
    Dim dic As IDictionary
    
    Set dic = Parser.ParseJSON(strBuf)
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

Sub BookReader_Sample()

    Dim BR As BookReader
    
    Set BR = BookReader.CreateObject("Sample.xlsx")
    
    
    Set BR = Nothing


End Sub




