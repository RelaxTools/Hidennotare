VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hecatoncheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hecatoncheir
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------
' フォーム１枚で実行中表示、キャンセル表示、コンボの履歴表示を行う画面用ワークフレーム
'
' コントロールのTagプロパティに設定した文字列により実行を制御する。
' e : 実行中でも活性化するコントロール
' g : ゲージの役割をするラベル
' m : メッセージ兼ゲージの背景の役割をするラベル
' c : 処理の中断を行うボタン。実行中は活性となりキャプションが「中断」になる。
'-----------------------------------------------------------------------------------------------------
' IUsing I/F にて 処理開始/終了を自動化可能。
'-----------------------------------------------------------------------------------------------------
Option Explicit

Implements IConstructor
Implements IUsing

Private mFrm As UserForm
Private mlngMax As Long
Private mstrBackup As String
Private mlngColor As Long
Private mblnCancel As Boolean
Private mblnRunning As Boolean

Private mobjGaugeLabel As MSForms.Label
Private mobjMessageLabel As MSForms.Label

Private mstrCancelStr As String
Private mstrBeforeName As String

Private Const C_CANCEL_STR As String = "中断"

Private mcolEnable As Collection
'--------------------------------------------------------------
' Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()

    mblnRunning = False
    mblnCancel = False
    
    '規定
    mstrCancelStr = C_CANCEL_STR
    
End Sub
'--------------------------------------------------------------
' Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()

    On Error Resume Next
    mFrm.lblBack.Caption = mstrBackup

End Sub
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
Private Function IConstructor_Instancing(Args As Collection) As Object
    
    Select Case Args.Count
        Case 1
            Set Me.Form = Args(1)
        Case 2
            Set Me.Form = Args(1)
            Me.ButtonCaption = Args(2)
        Case Else
    End Select
    
    Set IConstructor_Instancing = Me

End Function
'--------------------------------------------------------------
'  フォーム
'--------------------------------------------------------------
Property Set Form(frm As UserForm)

    Dim c As control
    
    Set mobjGaugeLabel = Nothing
    Set mobjMessageLabel = Nothing
    Set mFrm = frm

    For Each c In mFrm.Controls

        Select Case c.Tag
            Case "g"
                Set mobjGaugeLabel = c
                mobjGaugeLabel.Width = 0
                mobjGaugeLabel.ForeColor = vbWhite

            Case "m"
                Set mobjMessageLabel = c
                mstrBackup = mobjMessageLabel.Caption

        End Select

    Next

End Property
'--------------------------------------------------------------
'  中断ボタンの名称（中断以外にしたい場合設定）
'--------------------------------------------------------------
Public Property Let ButtonCaption(ByVal v As String)
    mstrCancelStr = v
End Property
'--------------------------------------------------------------
'  中断ボタンが押されたか？
'--------------------------------------------------------------
Public Property Get IsCancel() As Boolean
    IsCancel = mblnCancel
End Property
'--------------------------------------------------------------
'  実行中か？
'--------------------------------------------------------------
Public Property Get IsRunning() As Boolean
    IsRunning = mblnRunning
End Property
'--------------------------------------------------------------
'  処理開始
'--------------------------------------------------------------
Public Sub StartRunning()
    IUsing_Begin
End Sub
'--------------------------------------------------------------
'  処理終了
'--------------------------------------------------------------
Public Sub StopRunning()
    IUsing_Finish
End Sub
'--------------------------------------------------------------
'  処理開始(IUsing I/F)
'--------------------------------------------------------------
Private Sub IUsing_Begin()
    
    mblnCancel = False
    mblnRunning = True
    
    Set mcolEnable = New Collection

    If Not mFrm Is Nothing Then
        Dim c As control
        For Each c In mFrm.Controls
        
            Select Case c.Tag
                Case "e", "g", "m"
                    '何もしない
                Case "c"
                    mstrBeforeName = c.Caption
                    c.Caption = mstrCancelStr
                Case Else
                    'EnableになっているコントロールをDisableにする。ただし、元々Disableになっていたものは触らない
                    If c.Enabled Then
                        mcolEnable.Add c
                        c.Enabled = False
                    End If
            End Select
        
        Next
        DoEvents
    End If

End Sub
'--------------------------------------------------------------
'  処理終了(IUsing I/F)
'--------------------------------------------------------------
Private Sub IUsing_Finish()

    mblnRunning = False
    
    If Not mFrm Is Nothing Then
        
        Dim c As control
        For Each c In mFrm.Controls
            Select Case c.Tag
                Case "c"
                    c.Caption = mstrBeforeName
            End Select
        Next
        
        For Each c In mcolEnable
            c.Enabled = True
        Next
        Set mcolEnable = Nothing
    
    End If

    If Not mobjGaugeLabel Is Nothing Then
        mobjGaugeLabel.Caption = ""
        mobjGaugeLabel.Width = 0
    End If
    
    If Not mobjMessageLabel Is Nothing Then
        mobjMessageLabel.Caption = mstrBackup
    End If
    
End Sub

'--------------------------------------------------------------
'  ガイダンスメッセージ表示
'--------------------------------------------------------------
Public Sub DispGuidance(ByVal strValue As String)

    If Not mobjMessageLabel Is Nothing Then
        mobjMessageLabel.Caption = strValue
    End If
    
    'メッセージ表示
    DoEvents

End Sub
'--------------------------------------------------------------
'  進捗状況開始
'--------------------------------------------------------------
Public Sub StartGauge(ByVal lngValue As Long)
    
    If Not mobjGaugeLabel Is Nothing Then
        mlngMax = lngValue
    End If

End Sub
'--------------------------------------------------------------
'  進捗状況描画
'--------------------------------------------------------------
Public Sub DisplayGauge(ByVal lngValue As Long)

    Dim dblValue As Double
    Dim strMessage As String
    
    If Not mobjGaugeLabel Is Nothing Then
        
        If lngValue > mlngMax Then
            lngValue = mlngMax
        End If
        dblValue = (CDbl(lngValue) / mlngMax)
        mobjGaugeLabel.Width = mobjMessageLabel.Width * dblValue
        
        strMessage = Space$(Fix(mobjMessageLabel.Width * 0.16)) & Format$(Fix(dblValue * 100), "0") & "%"
        mobjMessageLabel.Caption = strMessage
        mobjGaugeLabel.Caption = strMessage
        
        'ゲージの表示
        DoEvents
    
    End If
    
End Sub
'--------------------------------------------------------------
'  キャンセルメソッド
'--------------------------------------------------------------
Public Sub doCancel()
    
    mblnCancel = True
    
    'ボタンの凹み表示用
    DoEvents

End Sub
'--------------------------------------------------------------
'  コンボボックスの履歴設定
'--------------------------------------------------------------
Public Function SetHistory(c As ComboBox, ByVal strHistory As String)

    Dim strFolder() As String
    Dim i As Long

    strFolder = Split(strHistory, vbTab)
    
    For i = LBound(strFolder) To UBound(strFolder)
        c.AddItem strFolder(i)
    Next
    If c.ListCount > 0 Then
        c.ListIndex = 0
    End If

End Function
'--------------------------------------------------------------
'  コンボボックスの履歴作成＆リスト更新
'--------------------------------------------------------------
Public Function CreateHistory(c As ComboBox, Optional ByVal max As Long = 10) As String
    
    Dim strBuf As String
    Dim i As Long
    Dim cnt As Long
    
    strBuf = c.Text
    
    '履歴の作成
    cnt = 1
    For i = 0 To c.ListCount - 1
        If c.list(i) <> c.Text Then
            strBuf = strBuf & vbTab & c.list(i)
            cnt = cnt + 1
            'リストは最大１０
            If cnt >= max Then
                Exit For
            End If
        End If
    Next
    
    '履歴を元に新しいリストを作成
    c.Clear
    
    Dim v As Variant
    
    v = Split(strBuf, vbTab)
    
    For i = LBound(v) To UBound(v)
        c.AddItem v(i)
    Next
    If c.ListCount > 0 Then
        c.ListIndex = 0
    End If
    
    CreateHistory = strBuf

End Function

