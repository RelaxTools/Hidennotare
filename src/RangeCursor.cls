VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeCursor"
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
' Range読み込みクラス
'-----------------------------------------------------------------------------------------------------
' Dim IC As ICursor
' Set IC = Constructor(New RangeCursor, Renge範囲)
' Do Until IC.Eof
'     Debug.Print IC(1).Value '範囲の１列目
'     IC.MoveNext
' Loop
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

'Range
Private mRange As Range

'下方向カーソル
Private mStartRow As Long
Private mEofCol As Variant
Private mRow As Long
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - Range   : UsedRange/ActiveSheet.Range(ActiveSheet.PrintArea) etc...
'--------------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Const RANGE_AREA As Long = 1  '引数１　RANGE オブジェクト
    
    Select Case Args.Count
        Case 1
        Case Else
            Exit Function
    End Select
    
    Set mRange = Args(RANGE_AREA)
                
    mRow = 1
    mStartRow = mRow
    
    SkipHiddenRow
    
    Set IConstructor_Instancing = Me

End Function
'------------------------------------------------------
' デストラクタ
'------------------------------------------------------
Private Sub Class_Terminate()
    Set mRange = Nothing
End Sub
'------------------------------------------------------
' 非表示行をスキップ
'------------------------------------------------------
Private Sub SkipHiddenRow()
    Do Until Not mRange.Rows(mRow).Hidden Or ICursor_Eof
        mRow = mRow + 1
    Loop
End Sub
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
    ICursor_Eof = (mRange.Rows.Count < mRow)
End Property
'------------------------------------------------------
' セル取得（規定のプロパティなので省略表記可）
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    Set ICursor_Item = mRange.Cells(mRow, opt)
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mRow = mStartRow
    SkipHiddenRow
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mRow = mRow + 1
    SkipHiddenRow
End Sub
