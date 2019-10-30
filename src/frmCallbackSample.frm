VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCallbackSample 
   Caption         =   "Callbackサンプル"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   OleObjectBlob   =   "frmCallbackSample.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCallbackSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Callbackクラスサンプル

Private WithEvents CB As Callback
Attribute CB.VB_VarHelpID = -1

'ユニークな番号を定義
Private Enum ActionConstants
    ActionMessageInfo = 0
    ActionMessageExcl
    DeleyExec
End Enum

Private Sub UserForm_Initialize()
    Set CB = New Callback
End Sub
Private Sub UserForm_Terminate()
    Set CB = Nothing
End Sub
' 右クリックメニュー表示
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    
    '右クリック
    If Button <> 2 Then
        Exit Sub
    End If
    
    'OnAction にCreateOnActionメソッドと番号を設定
    With CommandBars.Add(Position:=msoBarPopup, Temporary:=True)

        With .Controls.Add
            .BeginGroup = True
            .Caption = "情報メッセージ"
            .OnAction = CB.CreateOnAction(ActionMessageInfo)
            .FaceId = 535
        End With
        With .Controls.Add
            .Caption = "警告メッセージ"
            .OnAction = CB.CreateOnAction(ActionMessageExcl)
            .FaceId = 534
        End With
        
        .ShowPopup
    
    End With
    
End Sub
'OnTimeにも使える
Private Sub CommandButton1_Click()
    '３秒後に実行
    Process.UnsyncRun CB.CreateOnAction(DeleyExec), 3
End Sub

'実際の処理を記述するイベント
Private Sub CB_OnAction(ByVal Action As Long, ByVal opt As String)

    '実行された番号が戻ってくる
    Select Case Action
        Case ActionMessageInfo
            MsgBox "情報メッセージ", vbInformation
        Case ActionMessageExcl
            MsgBox "警告メッセージ", vbExclamation
        Case DeleyExec
            MsgBox "３秒経ちました", vbInformation
    End Select

End Sub

