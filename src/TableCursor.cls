VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableCursor"
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
' テーブル(ListObject)読み込みクラス
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

Private mLO As ListObject
Private mWS As Worksheet
Private mlngStartRow As Long
Private mlngRow As Long
Private mlngLastRow As Long
'------------------------------------------------------
' コンストラクタ
'------------------------------------------------------
Private Sub Class_Initialize()

End Sub
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    '引数：WorksheetまたはListObject
    Select Case Args.Count
        Case 1
            Select Case TypeName(Args(1))
                Case "Worksheet"
                    Set mWS = Args(1)
                    Set mLO = mWS.ListObjects(1)
                Case "ListObject"
                    Set mLO = Args(1)
                    Set mWS = mLO.Parent
                Case Else
                    Exit Function
            End Select
        Case Else
            Exit Function
    End Select
    
    If mLO.ListRows.Count > 0 Then
        'データの最初の行
        mlngStartRow = mLO.DataBodyRange(1).row
        'データの最後の行
        mlngLastRow = mLO.DataBodyRange(mLO.DataBodyRange.Count).row
        'カレント行
        mlngRow = mlngStartRow
        
        SkipHiddenRow
    Else
        mlngStartRow = 0
        mlngLastRow = 0
        mlngRow = 0
    End If
    
    Set IConstructor_Instancing = Me

End Function
'------------------------------------------------------
' デストラクタ
'------------------------------------------------------
Private Sub Class_Terminate()
    Set mLO = Nothing
    Set mWS = Nothing
End Sub
'------------------------------------------------------
' 非表示行をスキップ
'------------------------------------------------------
Private Sub SkipHiddenRow()
    Do Until Not mWS.Rows(mlngRow).Hidden Or ICursor_Eof
        mlngRow = mlngRow + 1
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
    ICursor_Eof = (mlngRow > mlngLastRow Or mlngRow < 1)
End Property
'------------------------------------------------------
' セル取得
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    Set ICursor_Item = mWS.Cells(mlngRow, mLO.ListColumns(opt).Range(1).Column)
End Property
'------------------------------------------------------
' 前のセル取得
'------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    Set ICursor_PreviousItem = mWS.Cells(mlngRow - 1, mLO.ListColumns(opt).Range(1).Column)
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mlngRow = mlngStartRow
    SkipHiddenRow
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mlngRow = mlngRow + 1
    SkipHiddenRow
End Sub
