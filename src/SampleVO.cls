VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'>### SampleVO クラス
'>
'>**Remarks**
'>
'>バリューオブジェクト(VO)サンプルクラス<br>
'>IValueObject I/F に対応したメソッドの書き方サンプル
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Implements IValueObject

Public Name As String
Public Age As Long
Public Address As String

Public mFields As Variant


Private Sub Class_Initialize()
    mFields = Array("Name", "Age", "Address")
End Sub

Private Property Let IValueObject_Fields(key As Variant, RHS As Variant)
    CallByName Me, key, VbLet, RHS
End Property

Private Property Get IValueObject_Fields(key As Variant) As Variant
    IValueObject_Fields = CallByName(Me, key, VbGet)
End Property

Private Property Get IValueObject_Count() As Long
    IValueObject_Count = Arrays.Length(mFields)
End Property

Private Property Get IValueObject_FieldNames() As Variant
    IValueObject_FieldNames = mFields
End Property

Private Function IValueObject_ToString() As String
    
    With StringBuilder.NewInstance
    
        .Append Name, True
        .Append CStr(Age)
        .Append Address, True
    
        IValueObject_ToString = .ToString(",", "{", "}")
    
    End With
    
End Function
