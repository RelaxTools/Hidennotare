VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hidennotare
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
'>### FormManager クラス
'>
'>**Remarks**
'>
'>フォーム管理クラス<br>
'>フォーム１枚で実行中表示、中断ボタン表示、コンボの履歴表示を行う画面用フレームワーク<br>
'>IUsing I/F にて 処理開始/終了を自動化可能。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>コントロールのTagプロパティに設定した文字列により実行を制御する。
'>
'>- e : 実行中でも活性化するコントロール
'>- g : ゲージの役割をするラベル
'>- m : メッセージ兼ゲージの背景の役割をするラベル
'>- c : 処理の中断を行うボタン。実行中は活性となりキャプションが「中断」になる。
'>
'>**Dependent**
'>
'>* IUsing
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

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
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = FormManager.CreateObject(
'>                    frm,
'>                    [ButtonCaption])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|frm|必須|UserForm|処理を行うUserFormを指定|
'>|ButtonCaption|省略可|中断ボタンのキャプションを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(FormManager)|
'>
'>**Remarks**
'>
'>**このクラスはコンストラクタを使用しないと実行できません。**
'>
Public Function CreateObject(frm As UserForm, Optional ByVal ButtonCaption As String = C_CANCEL_STR) As FormManager

    Dim FM As FormManager

    Set FM = New FormManager
    
    Set FM.Form = frm
    FM.ButtonCaption = ButtonCaption

    Set CreateObject = FM

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
'>---
'>#### IsCancel
'>
'>**Syntax**
'>
'>```
'>bool = expression.IsCancel()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|bool|Boolean|中断ボタンが押された:True/左記以外:False|
'>
'>**Remarks**
'>
'>中断ボタン(Tagに"c"が設定されたボタン）が<br>
'>押された際にTruenになります。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Property Get IsCancel() As Boolean
    IsCancel = mblnCancel
End Property
'>---
'>#### IsRunning
'>
'>**Syntax**
'>
'>```
'>bool = expression.IsRunning()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|bool|Boolean|処理が実行されている場合:True/左記以外:False|
'>
'>**Remarks**
'>
'>StartRunning/StopRunning または Using I/F で実行中になっているかどうかを<br>
'>確認するプロパティです。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Property Get IsRunning() As Boolean
    IsRunning = mblnRunning
End Property
'>---
'>#### StartRunning
'>
'>**Syntax**
'>
'>```
'>expression.StartRunning()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>Formでの処理の実行開始を宣言するメソッド。<br>
'>画面を非活性(Disable)j状態にします。<br>
'>Using I/F を使用すると簡略化が可能です。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub StartRunning()
    IUsing_Begin
End Sub
'>---
'>#### StopRunning
'>
'>**Syntax**
'>
'>```
'>expression.StopRunning()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>Formでの処理の実行終了を宣言するメソッド。<br>
'>画面を非活性(Disable)j状態を元に戻します。<br>
'>Using I/F を使用すると簡略化が可能です。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
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
'>---
'>#### DispGuidance
'>
'>**Syntax**
'>
'>```
'>expression.DispGuidance()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>|Value|必須|String|ガイダンスラベルに表示するメッセージを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ガイダンスラベルにメッセージを表示するメソッド<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub DispGuidance(ByVal strValue As String)

    If Not mobjMessageLabel Is Nothing Then
        mobjMessageLabel.Caption = strValue
    End If
    
    'メッセージ表示
    DoEvents

End Sub
'>---
'>#### StartGauge
'>
'>**Syntax**
'>
'>```
'>expression.StartGauge(Value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>|Value|必須|Long|処理を行う数の最大を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ゲージラベルの表示を開始します。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub StartGauge(ByVal lngValue As Long)
    
    If Not mobjGaugeLabel Is Nothing Then
        mlngMax = lngValue
    End If

End Sub
'>---
'>#### DisplayGauge
'>
'>**Syntax**
'>
'>```
'>expression.DisplayGauge(Value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>|Value|必須|Long|処理を行った数を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ゲージラベルの処理状況を更新します。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
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
'>---
'>#### doCancel
'>
'>**Syntax**
'>
'>```
'>expression.doCancel()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>中断ボタンの押下をFormMangerに通知します。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub doCancel()
    
    mblnCancel = True
    
    'ボタンの凹み表示用
    DoEvents

End Sub
'>---
'>#### SetHistory
'>
'>**Syntax**
'>
'>```
'>expression.SetHistory(obj, history)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>|obj|必須|ComboBox|履歴情報を設定するコンボボックス|
'>|history|必須|String|履歴情報をTab区切で設定（CreateHistoryの返却値）|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ComboBoxに履歴を表示する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
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
'>---
'>#### CreateHistory
'>
'>**Syntax**
'>
'>```
'>history = expression.CreateHistory(obj, max)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|FormManager|FormManager のインスタンス|
'>|obj|必須|ComboBox|履歴情報を設定するコンボボックス|
'>|max|省略可|Long|保存するコンボボックスの履歴数を指定。規定値は10|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|history|String|コンボボックスに表示されている文字をTab区切で返却|
'>
'>**Remarks**
'>
'>コンボボックスの内容を指定された数分、文字列に変換する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function CreateHistory(c As ComboBox, Optional ByVal Max As Long = 10) As String
    
    Dim strBuf As String
    Dim i As Long
    Dim cnt As Long
    
    strBuf = c.Text
    
    '履歴の作成
    cnt = 1
    For i = 0 To c.ListCount - 1
        If c.List(i) <> c.Text Then
            strBuf = strBuf & vbTab & c.List(i)
            cnt = cnt + 1
            'リストは最大１０
            If cnt >= Max Then
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

