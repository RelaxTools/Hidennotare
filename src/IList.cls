VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
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
'>### IList インターフェイス
'>
'>**Remarks**
'>
'>ArrayList等List系の動作を定義したインターフェイス
'>
'>**Definition**
'>
'>```
'>Public Sub Clear()
'>Public Sub Add(obj As Variant)
'>Public Sub Insert(idx As Long, v As Variant)
'>Public Function NewEnum() As stdole.IUnknown
'>    'Attribute NewEnum.VB_UserMemId = -4
'>Public Property Get Count() As Long
'>Public Property Get Item(idx As Long) As Variant
'>    'Attribute Item.VB_UserMemId = 0
'>Public Sub RemoveAt(idx As Long)
'>Public Sub Sort(Optional ByVal CP As IComparer = Nothing)
'>Public Function ToArray() As Variant
'>Public Function ToCollection() As Collection
'>Public Function ToString() As String
'>Public Property Get GetCursor() As ICursor
'>```
'>
'>**Sample**
'>```
'>Dim lst As IList
'>
'>Set lst = ArrayList.NewInstance()
'>
'>lst.Add "Data"
'>
'>```
'>
'-----------------------------------------------------------------------------------------------------
Public Sub Clear()
End Sub
Public Sub Add(obj As Variant)
End Sub
Public Sub Insert(idx As Long, v As Variant)
End Sub
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemId = -4
End Function
Public Property Get Count() As Long
End Property
Public Property Get Item(idx As Long) As Variant
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
End Property
Public Sub RemoveAt(idx As Long)
End Sub
Public Sub sort(Optional ByVal CP As IComparer = Nothing)
End Sub
Public Function ToArray() As Variant
End Function
Public Function ToCollection() As Collection
End Function
Public Function ToString() As String
End Function
Public Property Get GetCursor() As ICursor
End Property



