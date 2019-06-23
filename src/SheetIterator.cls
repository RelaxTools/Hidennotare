VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [HecatonCheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/HecatonCheir
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
' シート読み込みクラス(Iterator)
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements IIterator

'ワークシート
Private mWS As Worksheet

'下方向カーソル
Private mStartRow As Long
Private mEofCol As Variant

'右方向カーソル
Private mStartCol As Long
Private mEofRow As Variant

'保持用コレクション
Private mEnum As Collection

'EOF文字列
Private mEofStr As String

'カーソルの向き
Enum SheetIteratorConstants
    DirectionDown = 0           '下方向カーソル
    DirectionRight = 1          '右方向カーソル
End Enum
Private mDirection As SheetIteratorConstants

'------------------------------------------------------
' コンストラクタ
'------------------------------------------------------
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
        mDirection = SheetIteratorConstants.DirectionDown 'Default下方向
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
            
            If mDirection = SheetIteratorConstants.DirectionDown Then
                mStartRow = CLng(Args(START_NUM))
                mEofCol = Args(EOF_NUM)
            Else
                '指定が数値ならそのまま。アルファベットなら列変換
                If IsNumeric(Args(START_NUM)) Then
                    mStartCol = Args(START_NUM)
                Else
                    mStartCol = mWS.Columns(Args(START_NUM)).Column
                End If
                mEofRow = CLng(Args(EOF_NUM))
            End If
        
        Else
            Exit Function
        End If
    End If
    
    Set mEnum = New Collection

    Dim lngCnt As Long

    If mDirection = SheetIteratorConstants.DirectionDown Then
        lngCnt = mStartRow
        Do Until mWS.Cells(lngCnt, mEofCol).Value = mEofStr
            If Not mWS.Rows(lngCnt).Hidden Then
                mEnum.Add mWS.Rows(lngCnt).Columns
            End If
            lngCnt = lngCnt + 1
        Loop
    Else
        lngCnt = mStartCol
        Do Until mWS.Cells(mEofRow, lngCnt).Value = mEofStr
            If Not mWS.Columns(lngCnt).Hidden Then
                mEnum.Add mWS.Columns(lngCnt).Rows
            End If
            lngCnt = lngCnt + 1
        Loop
    End If
    
    Set IConstructor_Instancing = Me

End Function
'------------------------------------------------------
' デストラクタ
'------------------------------------------------------
Private Sub Class_Terminate()
    Set mEnum = Nothing
End Sub
'--------------------------------------------------------------
' IIterator インターフェースを取得
'--------------------------------------------------------------
Public Property Get GetIterator() As IIterator
    Set GetIterator = Me
End Property
'--------------------------------------------------------------
' For Each 用 IUnknown インタフェースの返却
'--------------------------------------------------------------
Private Function IIterator_NewEnum() As stdole.IEnumVARIANT
   Set IIterator_NewEnum = mEnum.[_NewEnum]
End Function
