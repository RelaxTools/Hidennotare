VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Callback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Callback クラス
'-----------------------------------------------------------------------------------------------------
' Application.Run / .OnKey / .OnAction / .OnUndo / .OnRepeat / .OnTime の呼び出しを
' クラスのイベントに変換するクラス。
'-----------------------------------------------------------------------------------------------------
Option Explicit

'メインイベント
Public Event OnAction(ByVal Action As Long, ByVal opt As String)

Public Tag As String
Private mKey As String
Public Function CreateOnAction(ByVal Action As Long, Optional ByVal opt As String = "") As String
    CreateOnAction = ExcelHelper.BuildOnAction("'CallbackHelper.OnActionCallback """ & mKey & """, """ & Action & """, """ & opt & """ '")
End Function
Public Sub OnActionCallback(ByVal Action As Long, ByVal opt As String)
    RaiseEvent OnAction(Action, opt)
End Sub
Private Sub Class_Initialize()
    mKey = CallbackHelper.InstallCallback(Me)
End Sub
Private Sub Class_Terminate()
    CallbackHelper.UninstallCallback mKey
End Sub
