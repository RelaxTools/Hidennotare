VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2.5
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 言語/プレースホルダ/エスケープ対応
'-----------------------------------------------------------------------------------------------------
' ブックのプロパティ[タイトル]を設定してください。
'-----------------------------------------------------------------------------------------------------
Option Explicit
Private mTitle As String
'---------------------------------------
' 初期化処理
'---------------------------------------
Private Sub Class_Initialize()
    mTitle = "Microsoft Excel"
End Sub
'>---
'>#### AppInfo(Message)
'>
'>**Syntax**
'>
'>```
'>Message.AppInfo = ThisWorkBook
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|title|必須|IAppInfo|アプリケーション情報I/F|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>メッセージボックスに表示するアプリケーション名を設定<br>
'>
'>**Example**
'>
'>```
'>Debug.Print Message.Information("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.Information
'>* Message.Error
'>* Message.Critical
'>* Message.Question
'>
Public Property Let AppInfo(v As IAppInfo)
    mTitle = v.AppName
End Property
'>---
'>#### Information(Message)
'>
'>**Syntax**
'>
'>```
'>Message.Information(message, parameter)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|プレースホルダを含むメッセージ|
'>|parameter|必須|String|置換文字列|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>情報メッセージ<br>
'>情報アイコンのメッセージを表示します。プレースホルダ{0}が使用可能です。<br>
'>
'>**Example**
'>
'>```
'>Debug.Print Message.Information("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.Error
'>* Message.Critical
'>* Message.Question
'>
Public Sub Information(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbInformation, mTitle
End Sub
'>---
'>#### Error(Message)
'>
'>**Syntax**
'>
'>```
'>Message.Error(message, parameter)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|プレースホルダを含むメッセージ|
'>|parameter|必須|String|置換文字列|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>警告メッセージ<br>
'>警告アイコンのメッセージを表示します。プレースホルダ{0}が使用可能です。<br>
'>
'>**Example**
'>
'>```
'>Debug.Print Message.Error("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.Information
'>* Message.Critical
'>* Message.Question
'>
Public Sub Error(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbExclamation, mTitle
End Sub
'>---
'>#### Critical(Message)
'>
'>**Syntax**
'>
'>```
'>Message.Critical(message, parameter)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|プレースホルダを含むメッセージ|
'>|parameter|必須|String|置換文字列|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>異常終了メッセージ<br>
'>異常終了アイコンのメッセージを表示します。プレースホルダ{0}が使用可能です。<br>
'>
'>**Example**
'>
'>```
'>Debug.Print Message.Critical("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.Information
'>* Message.Error
'>* Message.Question
'>
Public Sub Critical(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbCritical, mTitle
End Sub
'>---
'>#### Question(Message)
'>
'>**Syntax**
'>
'>```
'>Message.Question(message, parameter)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|プレースホルダを含むメッセージ|
'>|parameter|必須|String|置換文字列|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>ステータスメッセージ<br>
'>ステータスバーにメッセージを表示します。プレースホルダ{0}が使用可能です。<br>
'>
'>**Example**
'>
'>```
'>Message.Question("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.Information
'>* Message.Error
'>* Message.Critical
'>
Public Function Question(ByVal strMsg As String, ParamArray p()) As Boolean
    Question = VBA.MsgBox(PrivatePlaceHolder(strMsg, p), vbOKCancel + vbQuestion, mTitle) <> vbOK
End Function
'>---
'>#### StatusBar(Message)
'>
'>**Syntax**
'>
'>```
'>Message.StatusBar(message, parameter)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|プレースホルダを含むメッセージ|
'>|parameter|必須|String|置換文字列|
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>ステータスメッセージ<br>
'>ステータスバーにメッセージを表示します。プレースホルダ{0}が使用可能です。<br>
'>
'>**Example**
'>
'>```
'>Message.StatusBar("あなたは{0}座です。", "いて")
'>```
'>
'>**See also**
'>
'>* Message.StatusBarClear
'>
Public Sub StatusBar(ByVal strMsg As String, ParamArray p())
    Application.StatusBar = PrivatePlaceHolder(strMsg, p)
    DoEvents
End Sub
'>---
'>#### StatusBarClear(Message)
'>
'>**Syntax**
'>
'>```
'>Message.StatusBarClear()
'>```
'>
'>**Parameters**
'>
'>* なし
'>
'>**Return Value**
'>
'>* なし
'>
'>**Remarks**
'>
'>ステータスメッセージのクリア<br>
'>ステータスバーのメッセージをクリアします。<br>
'>
'>**See also**
'>
'>* Message.StatusBarClear
'>
Public Sub StatusBarClear()
    Application.StatusBar = False
    DoEvents
End Sub
'---------------------------------------
' プレースホルダ変換
'---------------------------------------
Private Function PrivatePlaceHolder(ByVal strMsg As String, ByVal p As Variant) As String

    strMsg = StringUtils.Unescape(strMsg)

    If UBound(p) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(p)
            strMsg = VBA.Replace(strMsg, "{" & CStr(i) & "}", p(i))
        Next
    End If

    PrivatePlaceHolder = strMsg

End Function

