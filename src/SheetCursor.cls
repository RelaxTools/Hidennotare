VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetCursor"
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
'>### SheetCursor クラス
'>
'>**Remarks**
'>
'>シート読み込みクラス
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>- ISheetCursor クラス
'>- ICursor クラス
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにNewInstance以外表示されないように修正|
'>
'>**Sample**
'>
'>```
'>Set IC = SheetCursor.NewInstance(ActiveSheet, 10, "B")
'>
'>Do Until IC.Eof
'>
'>    IC("C").Value = "XXX" 'Range
'>
'>    IC.MoveNext
'>Loop
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements ISheetCursorParam

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
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set return = SheetCursor.NewInstance(
'>                         worksheet,
'>                         StartRow/Column,
'>                         EOF Row/Column,
'>                         [Direction = DirectionDown],
'>                         [EOF String = ""])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Worksheet|必須|Worksheet|処理対象のワークシートを指定)|
'>|StartRow/Column|必須|Variant|処理対象データが最初の行または列を指定(列の場合にはアルファベットで指定可能)|
'>|EOF Row/Column|必須|Variant|処理対象の終了を判定する列または行を指定(列の場合にはアルファベットで指定可能)|
'>|Direction|省略可|SheetCursorConstants|処理を行う方向。下向き:DirectionDown/右向き:DirectionRightデフォルト:DirectionDown|
'>|EOF String|省略可|String|終了判定する文字列。デフォルト""|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|return|ICursor|新しいインスタンス(ICursor)|
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
Public Function NewInstance(ByRef WS As Worksheet, _
                            ByVal StartNum As Variant, _
                            ByVal EofNum As Variant, _
                            Optional ByVal Direction As SheetCursorConstants = SheetCursorConstants.DirectionDown, _
                            Optional ByVal EofStr As String = "") As ICursor

    Dim SC As ISheetCursorParam

    Set SC = New SheetCursor
    
    Set SC.Worksheet = WS
    SC.Direction = Direction
    SC.StartNum = StartNum
    SC.EofNum = EofNum

    Set NewInstance = SC
    
    NewInstance.MoveFirst

End Function
Private Property Let ISheetCursorParam_Direction(ByVal RHS As SheetCursorConstants)
    mDirection = RHS
End Property
Private Property Let ISheetCursorParam_EofNum(ByVal RHS As Variant)
    If mDirection = DirectionDown Then
        mEofCol = RHS
    Else
        mEofRow = CLng(RHS)
    End If
End Property
Private Property Let ISheetCursorParam_StartNum(ByVal RHS As Variant)
    If mDirection = DirectionDown Then
        mStartRow = RHS
    Else
        '指定が数値ならそのまま。アルファベットなら列変換
        If IsNumeric(RHS) Then
            mStartCol = RHS
        Else
            mStartCol = mWS.Columns(RHS).Column
        End If
    
    End If
End Property
Private Property Set ISheetCursorParam_Worksheet(RHS As Worksheet)
    Set mWS = RHS
End Property
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

'>---
'>#### Eof(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Eof()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Boolean|カーソルが終了:True / カーソルが終了していない:False|
'>
'>**Remarks**
'>
'>ICursor I/F のプロパティです。<br>
'>ICursor の Eof プロパティを取得。カーソルの終了を検出します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Property Get ICursor_Eof() As Boolean
    If mDirection = DirectionDown Then
        ICursor_Eof = (mWS.Cells(mRow, mEofCol).Value = mEofStr)
    Else
        ICursor_Eof = (mWS.Cells(mEofRow, mCol).Value = mEofStr)
    End If
End Property
'>---
'>#### Item(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Item()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|カーソル位置の内容を取得します。|
'>
'>**Remarks**
'>
'>ICursor I/F のプロパティです。<br>
'>ICursor の Item プロパティを取得。現在行の内容を取得します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    If mDirection = DirectionDown Then
        Set ICursor_Item = mWS.Cells(mRow, opt)
    Else
        Set ICursor_Item = mWS.Cells(opt, mCol)
    End If
End Property
'>---
'>#### MoveFirst(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>expression.MoveFirst()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ICursor I/F のメソッドです。<br>
'>カーソル位置の初期化を行います。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Sub ICursor_MoveFirst()
    If mDirection = DirectionDown Then
        mRow = mStartRow
        SkipHiddenRow
    Else
        mCol = mStartCol
        SkipHiddenCol
    End If
End Sub
'>---
'>#### MoveNext(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>expression.MoveNext()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ICursor I/F のメソッドです。<br>
'>カーソル位置を次に移動します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Sub ICursor_MoveNext()
    If mDirection = DirectionDown Then
        mRow = mRow + 1
        SkipHiddenRow
    Else
        mCol = mCol + 1
        SkipHiddenCol
    End If
End Sub

