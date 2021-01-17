VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormManagerSample 
   Caption         =   "UserForm1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445.001
   OleObjectBlob   =   "frmFormManagerSample.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFormManagerSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim FM As IFormManager

'FormManagerサンプル

Private Sub UserForm_Initialize()

'--------------------------------------------------------------------
    'わかりやすくコードで書いていますが、プロパティで
    '設定してしまってかまいません。
    
    'メッセージ及びプログレスバーの背景を表示するラベルのTagに"m"を設定する。
    lblBack.Tag = "m"
    
    'プログレスバーを表示するラベルのTagに"g"を設定する。
    lblGauge.Tag = "g"
    
    'キャンセルボタンのTagに"c"を設定する。
    cmdOk.Tag = "c"
    
    '実行中でも非活性にしないコントロールのTagに"e"を設定する。
    lblEnabled.Tag = "e"
    
'--------------------------------------------------------------------

    Set FM = FormManager.NewInstance(Me)

    FM.DispGuidance "開始しました。"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '実行中か？
    If FM.IsRunning Then
        'フォームが閉じられないようにする
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Set FM = Nothing
End Sub
Private Sub cmdOk_Click()

    '実行中か？実行中は中断ボタンになるので、キャンセル動作。
    If FM.IsRunning Then
        'キャンセルを実行する
        FM.doCancel
        Exit Sub
    End If

    If Message.Question("実行します。よろしいですか？") Then
        Exit Sub
    End If


    Dim lngMax As Long
    Dim i As Long
    
    '処理中を表す。以下メソッドを呼ぶかUsingクラスを使用する。
    'FM.StartRunning
    With Using.NewInstance(FM, New OneTimeSpeedBooster)
        
        lngMax = 10000
    
        'ゲージの最大を設定
        FM.StartGauge lngMax
        
        For i = 1 To lngMax
        
            'キャンセルか？
            If FM.IsCancel Then
                '処理を中断
                Exit For
            End If
        
        
            '処理を記述
        
        
            'ゲージの現在値を設定
            FM.DisplayGauge i
        Next
    
    End With
    '処理終了を表す
    'FM.StopRunning
    
    'キャンセルか？
    If FM.IsCancel Then
        Message.Error "処理は中断されました。"
    Else
        Message.Information "完了しました。"
    End If

End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

