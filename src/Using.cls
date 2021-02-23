VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Using"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'>### Using クラス
'>
'>**Remarks**
'>
'>開始／終了(With～End With)管理クラス<br>
'>コンストラクタに指定されたオブジェクトのIUsing I/F を呼び出す。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* IUsing
'>* IUsingParam
'>* IUsingArgs
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IUsingParam
Implements IUsingArgs

Private m_col As Collection
'--------------------------------
' With 時点で実行
'--------------------------------
Public Function NewInstance(ParamArray Args() As Variant) As IUsingArgs

    Dim us As IUsingParam
    Dim v As Variant
    
    v = Args
    
    Set us = New Using
    
    Set us.Source = Arrays.ToCollection(v)
    us.Init
    
    Set NewInstance = us

End Function
'>---
'>#### Args(IUsingArgs I/F)
'>
'>**Syntax**
'>
'>```
'>Set col = obj.Args()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|col|Collection|引数に渡されたオブジェクトのコレクション|
'>
'>**Remarks**
'>
'>CreteObjectに指定したオブジェクトを参照するためのコレクション<br>
'>Collectionなので1開始です。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Function IUsingArgs_Args(ByVal index As Long) As Object
    Set IUsingArgs_Args = m_col(index)
End Function

Private Property Set IUsingParam_Source(RHS As Collection)
    Set m_col = RHS
End Property
'--------------------------------
'With 時点で実行
'--------------------------------
Private Sub IUsingParam_Init()
    
    Dim v As IUsing
    
    For Each v In m_col
        v.Begin
    Next

End Sub
'--------------------------------
'End With 時点で実行
'--------------------------------
Private Sub Class_Terminate()

    Dim v As IUsing
    Dim i As Long
    
    'IUsingI/F同士に関連があるとアレなので、逆順に実行
    For i = m_col.Count To 1 Step -1
        Set v = m_col(i)
        v.Finish
    Next
    
    Set m_col = Nothing

End Sub



