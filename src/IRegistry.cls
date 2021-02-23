VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'>### IRegistry インターフェイス
'>
'>**Remarks**
'>
'>IRegistry のインターフェイス
'>
'>**Definition**
'>
'>```
'>Public Function GetSetting(ByVal Section As String, ByVal key As String, Optional ByVal Default As String = "") As String
'>Public Sub SaveSetting(ByVal Section As String, ByVal key As String, ByVal Value As String)
'>Public Sub DeleteSetting(ByVal Section As String, Optional ByVal key As String = "")
'>Public Sub Export()
'>```
'>
'>**Sample**
'>
'>* None
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------
' レジストリ読み込み
'--------------------------------------------------------------
Public Function GetSetting(ByVal Section As String, ByVal key As String, Optional ByVal Default As String = "") As String
End Function
'--------------------------------------------------------------
' レジストリ書き込み
'--------------------------------------------------------------
Public Sub SaveSetting(ByVal Section As String, ByVal key As String, ByVal Value As String)
End Sub
'--------------------------------------------------------------
'レジストリ削除
'--------------------------------------------------------------
Public Sub DeleteSetting(ByVal Section As String, Optional ByVal key As String = "")
End Sub
'--------------------------------------------------------------
' レジストリのExport
' UTF-16 LE にてレジストリの内容を出力する。
'--------------------------------------------------------------
Public Sub Export()
End Sub


