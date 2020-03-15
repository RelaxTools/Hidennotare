Attribute VB_Name = "CallbackHelper"
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
'>### CallbackHelper 標準モジュール
'>
'>**Remarks**
'>
'>- Callbackクラスと連携してクラスモジュール内にOnAction等のロジックをカプセル化する。
'>
'-----------------------------------------------------------------------------------------------------
Option Private Module
Option Explicit

'Callback用
Private mCallback As IDictionary
'---------------------------------------------------------------------------------------------------
'　Callbackの際のInstallメソッド
'---------------------------------------------------------------------------------------------------
Public Function InstallCallback(MH As Callback) As String

    Dim key As String

    If mCallback Is Nothing Then
        Set mCallback = Dictionary.NewInstance
    End If
    
    key = CStr(ObjPtr(MH))
    
    mCallback.Add key, MH
    
    InstallCallback = key
    
End Function
'---------------------------------------------------------------------------------------------------
'　Callbackの際のUnInstallメソッド
'---------------------------------------------------------------------------------------------------
Public Sub UninstallCallback(ByVal key As String)

    If mCallback.ContainsKey(key) Then
        mCallback.Remove key
    End If
    
End Sub
'---------------------------------------------------------------------------------------------------
'　Callbackの際に呼び出されるメソッド
'---------------------------------------------------------------------------------------------------
Public Function OnActionCallback(ByVal key As String, ByVal lngEvent As Long, ByVal opt As String)

    Dim MH As Callback
    
    If mCallback.ContainsKey(key) Then
        Set MH = mCallback(key)
        Call MH.OnActionCallback(lngEvent, opt)
    End If

End Function
