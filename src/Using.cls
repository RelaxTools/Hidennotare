VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Using"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2
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
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private m_col As Collection
'--------------------------------
' With 時点で実行
'--------------------------------
Public Function CreateObject(ParamArray Args() As Variant) As Using

    Dim us As Using
    Dim v As Variant
    
    v = Args
    
    Set us = New Using
    
    Set us.Source = Core.ArrayToCollection(v)
    us.Init
    
    Set CreateObject = us

End Function
Public Property Set Source(v As Collection)
    Set m_col = v
End Property
Public Sub Init()
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
'--------------------------------
' Args
'--------------------------------
Public Property Get Args() As Collection
    Set Args = m_col
End Property


