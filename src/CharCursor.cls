VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharCursor"
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
'>### CharCursor クラス
'>
'>**Remarks**
'>
'>カーソル風文字列読み込みクラス<br>
'>指定された文字列から１文字ずつ処理を行う。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>- ICharCursorParam クラス
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
'>Dim strBuf As String
'>strBuf ="123"
'>
'>Dim IC As ICursor
'>Set IC = CharCursor.NewInstance(strBuf)
'>
'>Do Until IC.Eof
'>
'>    Debug.Print IC.Item '1文字
'>
'>    IC.MoveNext
'>
'>Loop
'>```
'>
'>**Immediate**
'>```
'>1
'>2
'>3
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Implements ICursor
Implements IParam

Private mBuf As String
Private mIndex As Long
Private Sub Class_Initialize()
    mIndex = 1
End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = CharCursor.NewInstance(Source)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Source|必須|String|１文字ずつ処理する文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(ICursor)|
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
Public Function NewInstance(ByVal Source As String) As ICursor

    Dim p As IParam
    
    Set p = New CharCursor
    
    p.Parameters = Source
    
    Set NewInstance = p

    NewInstance.MoveFirst

End Function
Private Property Let IParam_Parameters(ByVal RHS As Variant)

    mBuf = RHS

End Property
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
    ICursor_Eof = Len(mBuf) < mIndex
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
    mIndex = 1
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
    mIndex = mIndex + 1
End Sub
'>---
'>#### Item(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Item([opt = 0])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>|opt|省略可|Long|相対位置を数値で指定する。0:現在文字/1:次文字/-1:前文字/省略時:0|
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

    Dim lngPos As Long
    
    If IsMissing(opt) Then
        opt = 0
    End If
    
    lngPos = mIndex + opt

    If lngPos < 1 Or lngPos > Len(mBuf) Then
        ICursor_Item = ""
    Else
        ICursor_Item = Mid$(mBuf, lngPos, 1)
    End If

End Property

