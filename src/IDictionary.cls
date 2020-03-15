VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDictionary"
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
'>### IDictionary インターフェイス
'>
'>**Remarks**
'>
'>Dictionary, OrderedDictionary, SortedDictionary等 Dictionary 系の動作を定義したインターフェイス
'>
'>**Definition**
'>
'>```
'>Public Sub Add(Key As Variant, Item As Variant)
'>Public Function NewEnum() As stdole.IUnknown
'>Public Property Get Count() As Long
'>Public Function ContainsKey(Key As Variant) As Boolean
'>Public Property Let Item(Key As Variant, RHS As Variant)
'>    'Attribute Item.VB_UserMemId = 0
'>Public Property Get Item(Key As Variant) As Variant
'>Public Property Set Item(Key As Variant, RHS As Variant)
'>Public Property Let Key(Key As Variant, RHS As Variant)
'>Public Function Keys() As Variant
'>Public Sub Remove(Key As Variant)
'>Public Sub Clear()
'>Public Function ToString() As String
'>Public Property Get GetCursor() As ICursor
'>```
'>
'>**Sample**
'>```
'>Dim dic As IDictionary
'>
'>Set dic = Dictionary.NewInstance
'>
'>dic.Add Key, Data
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Sub Add(key As Variant, Item As Variant)

End Sub
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_UserMemId = -4
End Function
Public Property Get Count() As Long

End Property

Public Function ContainsKey(key As Variant) As Boolean

End Function

Public Property Let Item(key As Variant, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
End Property

Public Property Get Item(key As Variant) As Variant

End Property

Public Property Set Item(key As Variant, RHS As Variant)

End Property

Public Property Let key(key As Variant, RHS As Variant)

End Property

Public Function Keys() As Variant

End Function

Public Sub Remove(key As Variant)

End Sub

Public Sub Clear()

End Sub
Public Function ToString() As String

End Function

Public Property Get GetCursor() As ICursor

End Property

Public Property Set Comparer(v As IComparer)

End Property
