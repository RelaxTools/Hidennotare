VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineCursor"
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
' カーソル風配列読み込みクラス
'-----------------------------------------------------------------------------------------------------
' Dim IC As ICursor
' Set IC = Constructor(New LineCursor, 配列)
'
' Do Until IC.Eof
'
'     Debug.Print IC.Item '配列の要素
'
'     IC.MoveNext
' Loop
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

Private mBuf As Variant
Private mIndex As Long
'------------------------------------------------------
' コンストラクタ
'------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Select Case Args.Count
        Case 1
            If VarType(Args(1)) = vbArray Then
                mBuf = Args(1)
            Else
                Exit Function
            End If
        Case Else
            Exit Function
    End Select
    
    ICursor_MoveFirst
    Set IConstructor_Instancing = Me
    
End Function
'--------------------------------------------------------------
' ICursor インターフェースを取得
'--------------------------------------------------------------
Public Property Get GetCursor() As ICursor
    Set GetCursor = Me
End Property
'------------------------------------------------------
' 終了判定
'------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = UBound(mBuf) < mIndex
End Property
'------------------------------------------------------
' 現在の行取得
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    If ICursor_Eof Then
        ICursor_Item = mBuf(UBound(mBuf))
    Else
        ICursor_Item = mBuf(mIndex)
    End If
End Property
'------------------------------------------------------
' 前の行取得
'------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    If ICursor_Eof Then
        ICursor_PreviousItem = mBuf(UBound(mBuf))
    Else
        ICursor_PreviousItem = mBuf(mIndex - 1)
    End If
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = LBound(mBuf)
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mIndex = mIndex + 1
End Sub

