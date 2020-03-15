VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableCursor"
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
'>### TableCursor クラス
'>
'>**Remarks**
'>
'>テーブル(ListObject)読み込みクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>- ITableCursorParam クラス
'>- ICursor クラス
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにCreateObject以外表示されないように修正|
'>
'>**Sample**
'>
'>```
'>Dim IC Ad ICursor
'
'>Set IC = TableCursor.NewInstance(WS.ListObjects(1))
'
'>Do Until IC.Eof
'
'>    IC("Item").Value = "XXX" '列ヘッダ名が使える
'
'>    IC.MoveNext
'>Loop
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements ITableCursorParam

Private mLO As ListObject
Private mWS As Worksheet
Private mlngStartRow As Long
Private mlngRow As Long
Private mlngLastRow As Long
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - Worksheet or Listobject
'--------------------------------------------------------------
Private Sub Class_Initialize()

End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set return = TableCursor.NewInstance(listobject)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|listobject|必須|Listobject|処理対象のListObject(テーブル)を指定)|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|return|Object|新しいインスタンス(ICursor)|
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
Public Function NewInstance(LO As ListObject) As ICursor
    
    Dim TC As ITableCursorParam
    
    Set TC = New TableCursor
    
    Set TC.ListObject = LO
    
    Set NewInstance = TC
    
    NewInstance.MoveFirst
    
End Function
Private Property Set ITableCursorParam_ListObject(RHS As ListObject)
    
    Set mLO = RHS
    Set mWS = mLO.Parent

End Property
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
    ICursor_Eof = (mlngRow > mlngLastRow Or mlngRow < 1)
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
    Set ICursor_Item = mWS.Cells(mlngRow, mLO.ListColumns(opt).Range(1).Column)
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
    
    If mLO.ListRows.Count > 0 Then
        'データの最初の行
        mlngStartRow = mLO.DataBodyRange(1).Row
        'データの最後の行
        mlngLastRow = mLO.DataBodyRange(mLO.DataBodyRange.Count).Row
        'カレント行
        mlngRow = mlngStartRow
        
        SkipHiddenRow
    Else
        mlngStartRow = 0
        mlngLastRow = 0
        mlngRow = 0
    End If
    
    mlngRow = mlngStartRow
    SkipHiddenRow
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
    mlngRow = mlngRow + 1
    SkipHiddenRow
End Sub


