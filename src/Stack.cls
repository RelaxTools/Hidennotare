VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [HecatonCheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/HecatonCheir
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
Option Explicit
Dim mCol As Collection
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'---------------------
' ï€ë∂(Push)
'---------------------
Sub Push(ByVal strBuf As String)
    mCol.Add strBuf
End Sub
'---------------------
' éÊÇËèoÇµ(Pop)
'---------------------
Function Pop() As String
    If mCol.Count = 0 Then
        Pop = ""
    Else
        Pop = mCol(mCol.Count)
        mCol.Remove mCol.Count
    End If
End Function
'---------------------
' éQè∆(çÌèúÇ»Çµ)
'---------------------
Function Peek() As String
    If mCol.Count = 0 Then
        Peek = ""
    Else
        Peek = mCol(mCol.Count)
    End If
End Function
'---------------------
' å¬êî
'---------------------
Property Get Count() As Long
    Count = mCol.Count
End Property
'Function Eof() As Boolean
'    If mCol.Count = 0 Then
'        Eof = True
'    Else
'        Eof = False
'    End If
'End Function
