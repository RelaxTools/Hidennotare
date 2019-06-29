VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 言語/プレースホルダ/エスケープ対応
'-----------------------------------------------------------------------------------------------------
' ブックのプロパティ[タイトル]を設定してください。
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private mLang As Object
'---------------------------------------
' 言語サポート
'---------------------------------------
Public Property Set LangSupport(v As Object)
    Set mLang = v
End Property
'---------------------------------------
' 情報メッセージ
'---------------------------------------
Public Sub Information(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbInformation, Title
End Sub
'---------------------------------------
' 警告メッセージ
'---------------------------------------
Public Sub Error(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbExclamation, Title
End Sub
'---------------------------------------
' 異常終了メッセージ
'---------------------------------------
Public Sub Critical(ByVal strMsg As String, ParamArray p())
    VBA.MsgBox PrivatePlaceHolder(strMsg, p), vbOKOnly + vbCritical, Title
End Sub
'---------------------------------------
' 応答メッセージ
'---------------------------------------
Public Function Question(ByVal strMsg As String, ParamArray p()) As Boolean
    Question = VBA.MsgBox(PrivatePlaceHolder(strMsg, p), vbOKCancel + vbQuestion, Title) <> vbOK
End Function
'---------------------------------------
' プレースホルダ変換
'---------------------------------------
Public Function PlaceHolder(ByVal strMsg As String, ParamArray p()) As String
    PlaceHolder = PrivatePlaceHolder(strMsg, p)
End Function
'---------------------------------------
' ステータスバーメッセージ
'---------------------------------------
Public Sub StatusBar(ByVal strMsg As String, ParamArray p())
    Application.StatusBar = PrivatePlaceHolder(strMsg, p)
    DoEvents
End Sub
'---------------------------------------
' ステータスバークリア
'---------------------------------------
Public Sub StatusBarClear()
    Application.StatusBar = False
    DoEvents
End Sub
'---------------------------------------
' プレースホルダ変換
'---------------------------------------
Private Function PrivatePlaceHolder(ByVal strMsg As String, ByVal p As Variant) As String
    
    '翻訳
    If Not mLang Is Nothing Then
        strMsg = mLang.Trunslate(strMsg)
    Else
        strMsg = Convert.UnEscape(strMsg)
    End If
    
    If UBound(p) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(p)
            strMsg = Replace(strMsg, "{" & CStr(i) & "}", p(i))
        Next
    End If

    PrivatePlaceHolder = strMsg

End Function
'---------------------------------------
' エラー処理
'---------------------------------------
Public Sub Throw(ByVal lngNo As Long, ByVal Module As Variant, ByVal strSrc As String, ByVal strMsg As String, ParamArray p())

    Dim strMessage As String
    Dim strSource As String
    
    If IsObject(Module) Then
        strSource = TypeName(Module) & "." & strSrc
    Else
        strSource = CStr(Module) & "." & strSrc
    End If

    'プレースホルダ変換
    strMessage = PrivatePlaceHolder(strMsg, p)
    
    Err.Raise 512 + lngNo, strSource, strMessage

End Sub
' アプリケーション名取得
Private Property Get Title() As String
    Title = ThisWorkbook.BuiltinDocumentProperties("Title").Value
End Property


