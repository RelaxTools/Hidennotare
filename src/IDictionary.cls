VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDictionary"
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
Option Explicit

'Public Enum CompareMethod
'    BinaryCompare = 0
'    TextCompare = 1
'    DatabaseCompare = 2
'End Enum
Public Function Add(ByVal KeyValue As String, obj As Variant)
End Function
Public Function ContainsKey(ByVal KeyValue As String) As Boolean
End Function
Public Property Get Keys() As Variant
End Property
Public Property Get Count() As Long
End Property
Public Property Get Item(ByVal KeyValue As String) As Variant
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
End Property
Public Property Let Item(ByVal KeyValue As String, NewItem As Variant)
End Property
Public Property Set Item(ByVal KeyValue As String, NewItem As Variant)
End Property
Public Property Let Key(ByVal OldKey As String, NewKey As Variant)
End Property
Public Sub Remove(ByVal KeyValue As String)
End Sub
Public Sub Clear()
End Sub
Public Function ToString() As String
End Function
Public Property Get GetCursor() As ICursor
End Property




