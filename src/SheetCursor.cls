VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetCursor"
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
' シート読み込みクラス(Cursor)
'-----------------------------------------------------------------------------------------------------
' Dim IC Ad ICursor
'
' Set IC = Core.Constructor(New SheetCursor, ActiveSheet, 10, "B")
'
' Do Until IC.Eof
'
'     IC("C").Value = "XXX" 'Range
'
'     IC.MoveNext
' Loop
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

'ワークシート
Private mWS As Worksheet

'下方向カーソル
Private mStartRow As Long
Private mEofCol As Variant
Private mRow As Long

'右方向カーソル
Private mStartCol As Variant
Private mEofRow As Long
Private mCol As Variant

'EOF文字列
Private mEofStr As String

'カーソルの向き
Enum SheetCursorConstants
    DirectionDown = 0           '下方向カーソル
    DirectionRight = 1          '右方向カーソル
End Enum
Private mDirection As SheetCursorConstants
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - Worksheet
' Arg2 - StartRow/Column : Start Row or Column
' Arg3 - EOF Row/Column  : Eof Row or Column
' Arg4 - Direction       : SheetCursorConstants.DirectionDown/DirectionRight
' Arg5 - EOF String      : Default ""
'--------------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Const WORK_SHEET As Long = 1  '引数１　Worksheet オブジェクト
    Const START_NUM As Long = 2   '引数２　開始行 or 列
    Const EOF_NUM As Long = 3     '引数３　終了を判定する行 or 列
    Const DIRECTION As Long = 4   '引数４　検索方向:下 or 右(default:下)
    Const EOF_STR As Long = 5     '引数５　終了を判定する文字列(default:""）
    
    '引数の数が３～５以外はエラー
    Select Case Args.Count
        Case 3 To 5
        Case Else
            Exit Function
    End Select
    
    'カーソル方向の指定
    If Args.Count >= DIRECTION Then
        mDirection = Args(DIRECTION)
    Else
        mDirection = SheetCursorConstants.DirectionDown 'Default下方向
    End If
    
    'EOF文字列の指定
    If Args.Count >= EOF_STR Then
        mEofStr = Args(EOF_STR)
    Else
        mEofStr = "" 'Default空文字列
    End If
    
    If Args.Count >= EOF_NUM Then
        
        If TypeOf Args(WORK_SHEET) Is Worksheet Then
            
            Set mWS = Args(WORK_SHEET)
            
            If mDirection = DirectionDown Then
                mRow = CLng(Args(START_NUM))
                mStartRow = mRow
                mEofCol = Args(EOF_NUM)
            Else
                '指定が数値ならそのまま。アルファベットなら列変換
                If IsNumeric(Args(START_NUM)) Then
                    mCol = Args(START_NUM)
                Else
                    mCol = mWS.Columns(Args(START_NUM)).Column
                End If
                
                mStartCol = mCol
                
                mEofRow = CLng(Args(EOF_NUM))
            
            End If
        Else
            Exit Function
        End If
    End If
    
    If mDirection = DirectionDown Then
        SkipHiddenRow
    Else
        SkipHiddenCol
    End If
    
    Set IConstructor_Instancing = Me

End Function
'------------------------------------------------------
' デストラクタ
'------------------------------------------------------
Private Sub Class_Terminate()
    Set mWS = Nothing
End Sub
'------------------------------------------------------
' 非表示行をスキップ
'------------------------------------------------------
Private Sub SkipHiddenRow()
    Do Until Not mWS.Rows(mRow).Hidden Or ICursor_Eof
        mRow = mRow + 1
    Loop
End Sub
'------------------------------------------------------
' 非表示列をスキップ
'------------------------------------------------------
Private Sub SkipHiddenCol()
    Do Until Not mWS.Columns(mCol).Hidden Or ICursor_Eof
        mCol = mCol + 1
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
    If mDirection = DirectionDown Then
        ICursor_Eof = (mWS.Cells(mRow, mEofCol).Value = mEofStr)
    Else
        ICursor_Eof = (mWS.Cells(mEofRow, mCol).Value = mEofStr)
    End If
End Property
'------------------------------------------------------
' セル取得（規定のプロパティなので省略表記可）
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    If mDirection = DirectionDown Then
        Set ICursor_Item = mWS.Cells(mRow, opt)
    Else
        Set ICursor_Item = mWS.Cells(opt, mCol)
    End If
End Property
'------------------------------------------------------
' 前のセル取得
'------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    If mDirection = DirectionDown Then
        Set ICursor_PreviousItem = mWS.Cells(mRow - 1, opt)
    Else
        Set ICursor_PreviousItem = mWS.Cells(opt, mCol - 1)
    End If
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    If mDirection = DirectionDown Then
        mRow = mStartRow
        SkipHiddenRow
    Else
        mCol = mStartCol
        SkipHiddenCol
    End If
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    If mDirection = DirectionDown Then
        mRow = mRow + 1
        SkipHiddenRow
    Else
        mCol = mCol + 1
        SkipHiddenCol
    End If
End Sub
