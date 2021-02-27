VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
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
'>### StringBuilder クラス
'>
'>**Remarks**
'>
'>文字列連結クラス
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Implements**
'>
'>* なし
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'>**Sample**
'>
'>```
'>Dim SB As IStringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Append "C"
'>
'>Debug.Print SB.ToString
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IStringBuilder
Implements IParam

Private mstrBuf() As String
Private mlngCount As Long
Private mlngLength As Long
Private mlngInitCount As Long
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set SB = StringBuilder.NewInstance(Number)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Number|省略可|Long|初期配列サイズを指定。デフォルト25|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|SB|Object|新しいインスタンス(StringBuilder)|
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>初期の配列数を指定。大量データの場合に事前に指定すると効率が良い<br>
'>
'>**Example**
'>
'>```
'>Dim SB As IStringBuilder
'>
'>Set SB = StringBuilder.NewInstance(3000)
'>
'>For i = 1 To 3000
'>
'>    SB.Append i
'>
'>Next
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function NewInstance(Optional ByVal size As Long = 25) As IStringBuilder

    Dim p As IParam
    
    Set p = New StringBuilder

    p.Parameters = size
    
    Set NewInstance = p
    
    NewInstance.Clear

End Function

Private Sub Class_Initialize()
    Call IStringBuilder_Clear
End Sub

Private Property Let IParam_Parameters(ByVal RHS As Variant)
    mlngInitCount = RHS
End Property
'------------------------------------------------------
' Terminate
'------------------------------------------------------
Private Sub Class_Terminate()
    Erase mstrBuf
End Sub

'>---
'>#### Clear(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Clear()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>クリアメソッド<br>
'>文字列バッファをクリアします。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Sub IStringBuilder_Clear()
    
    Erase mstrBuf
    ReDim Preserve mstrBuf(0 To mlngInitCount)
    
    mlngCount = 0
    mlngLength = 0

End Sub
'>---
'>#### Append(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Append(strBuf, [quotation = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>|strBuf|必須|String|連結を行う文字列を指定します。|
'>|quotation|省略可|Boolean|Trueの場合、ダブルコーテーションを付加する。デフォルトFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>文字列を追加するメソッドです。<br>
'>第２引数をTrueにするとダブルコーテーションで囲みます。
'>
'>**Example**
'>
'>半角0～9以外の文字列だったらダブルコーテーションで囲みます。
'>
'>```
'>Dim SB As StringBuilder
'>Dim strBuf As String
'>
'>Set SB = StringBuilder.NewInstance
'>
'>strBuf = "A"
'>
'>SB.Append strBuf, Not RegExp.Test(strBuf, "[0-9]+")
'>Debug.Print SB.ToString
'>
'>```
'>
'>結果
'>
'>```
'>"A"
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Private Sub IStringBuilder_Append(ByVal s As String, Optional ByVal Quotation As Boolean = False)

    '配列にセット
    If mlngCount > UBound(mstrBuf) Then
        ReDim Preserve mstrBuf(0 To mlngCount)
    End If
    
    If Quotation Then
        mstrBuf(mlngCount) = """" & s & """"
    Else
        mstrBuf(mlngCount) = s
    End If
    
    '文字数をカウント
    mlngLength = mlngLength + Len(mstrBuf(mlngCount))
    
    '要素数＋１
    mlngCount = mlngCount + 1

End Sub
'>---
'>#### ToString(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToString([Delimiter = "", Header = "", Footer= ""])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>|Delimiter|省略可|String|連結を行う際のデリミタを指定しますデフォルト""|
'>|Header|省略可|String|最初に連結する文字列を指定しますデフォルト""|
'>|Footer|省略可|String|連最後に連結する文字列を指定しますデフォルト""|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|連結後の文字列|
'>
'>**Remarks**
'>
'>文字列変換メソッドです。<br>
'>デリミタ、ヘッダ、フッタを指定可能です。<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Append "C"
'>
'>Debug.Print SB.ToString
'>
'>```
'>
'>結果
'>
'>```
'>ABC
'>
'>```
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Append "C"
'>
'>Debug.Print SB.ToString(",","[","]")
'>
'>```
'>
'>結果
'>
'>```
'>[A,B,C]
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Private Function IStringBuilder_ToString(Optional ByVal Delimiter As String = "", Optional ByVal Header As String = "", Optional ByVal Footer As String = "") As String

    Call resize
    IStringBuilder_ToString = Header & Join(mstrBuf, Delimiter) & Footer

End Function
'>---
'>#### ToArray(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToArray()
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
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>|var|String|変換後の配列|
'>
'>**Remarks**
'>
'>配列変換メソッド<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Append "C"
'>
'>Debug.Print SB.ToArray(1)
'>
'>```
'>
'>結果
'>
'>```
'>B
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Private Function IStringBuilder_ToArray() As String()
    
    Call resize
    IStringBuilder_ToArray = mstrBuf()

End Function
'>---
'>#### Length(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Length()
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
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>|var|Long|文字数を返却|
'>
'>**Remarks**
'>
'>文字数取得プロパティ<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Append "C"
'>
'>Debug.Print SB.Length
'>
'>```
'>
'>結果
'>
'>```
'>3
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Private Property Get IStringBuilder_Length() As Long
    IStringBuilder_Length = mlngLength
End Property
'>---
'>#### Replace(IStringBuilder I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Replace(strFind, strReplace)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|StringBuilderオブジェクトを指定します。|
'>|strFind|必須|String|置換前の文字列|
'>|strReplace|必須|String|置換後の文字列|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>置換メソッド<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = StringBuilder.NewInstance
'>
'>SB.Append "A"
'>SB.Append "B"
'>SB.Replace "B", "C"
'>
'>Debug.Print SB.ToString
'>
'>```
'>
'>結果
'>
'>```
'>AC
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Private Sub IStringBuilder_Replace(ByVal strFind As String, ByVal strReplace As String)

    Dim i As Long
    If mlngCount > 0 Then
        For i = 0 To mlngCount - 1
            mstrBuf(i) = VBA.Replace(mstrBuf(i), strFind, strReplace)
        Next
    End If
    
End Sub
'------------------------------------------------------
' 文字列変換前のリサイズ
'------------------------------------------------------
Private Sub resize()
    
    Select Case mlngCount
        Case Is <= 0
            ReDim Preserve mstrBuf(0 To 0)
        Case Is < mlngInitCount
            ReDim Preserve mstrBuf(0 To mlngCount - 1)
    End Select

End Sub

