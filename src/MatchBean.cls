VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v1
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
'>### MatchBean クラス
'>
'>**Remarks**
'>
'>Executeメソッド結果保持クラス
'>RegExpクラスとStdSchクラスのExecuteメソッドの戻り値を保持します。<br>
'>
'>**Dependent**
'>
'>* IConstructor
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor

Private m_index As Long
Private m_Length As Long
Private m_Value As String
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = Constructor(New MatchBean,
'>                      index,
'>                      length,
'>                      value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|index|必須|Long|開始場所を指定|
'>|length|必須|Long|文字数を指定|
'>|value|必須|String|検索結果を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(MatchBean)|
'>
'>**Remarks**
'>
'>** このクラスはコンストラクタを使用しないと実行できません。**
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
'>
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    If Args.Count = 3 Then
        
        m_index = Args(1)
        m_Length = Args(2)
        m_Value = Args(3)
        
        Set IConstructor_Instancing = Me
    
    End If
    
End Function
Public Property Get index() As Long
    index = m_index
End Property
Public Property Get Length() As Long
    Length = m_Length
End Property
Public Property Get Value() As String
    Value = m_Value
End Property

