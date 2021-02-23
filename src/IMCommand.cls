VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMCommand"
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
'>### IMCommand インターフェイス
'>
'>**Remarks**
'>
'>IMCommand のインターフェイス
'>
'>**Definition**
'>
'>```
'>Public Property Let AliasName(v As String)
'>Public Sub Append(m As MTable, Optional ByVal optName As String = "")
'>Public Property Get Table() As MTable
'>Public Function ToString()
'>```
'>
'>**Sample**
'>
'>* None
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Property Let AliasName(v As String)
End Property
Public Sub Append(m As MTable, Optional ByVal optName As String = "")
End Sub
Public Property Get Table() As MTable
End Property
Public Function ToString()
End Function


