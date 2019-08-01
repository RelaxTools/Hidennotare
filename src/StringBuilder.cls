VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v1
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
'>
'>**Implements**
'>
'>* IConstructor
'>
'>**Sample**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = New StringBuilder
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
Implements IConstructor

Private mstrBuf() As String
Private mlngCount As Long
Private mlngLength As Long
Private Const C_INIT_COUNT As Long = 25
Private mlngInitCount As Long
'------------------------------------------------------
' Initialize
'------------------------------------------------------
Private Sub Class_Initialize()
    mlngInitCount = C_INIT_COUNT
    Me.Clear
End Sub
'------------------------------------------------------
' Terminate
'------------------------------------------------------
Private Sub Class_Terminate()
    Erase mstrBuf
End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set IC = Constructor(New StringBuilder, Number)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Number|必須|Long|初期配列サイズを指定。デフォルト25|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|IC|Object|新しいインスタンス(StringBuilder)|
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>初期の配列数を指定。大量データの場合に事前に指定すると効率が良い<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>
'>Set SB = Constructor(New StringBuilder, 3000)
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
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    If Args.Count = 1 Then
        mlngInitCount = Args(1)
        Me.Clear
    End If
    
    Set IConstructor_Instancing = Me
    
End Function
'>---
'>#### Clear(StringBuilder)
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
'>なし
'>
'>**Remarks**
'>
'>クリアメソッド<br>
'>文字列バッファをクリアします。<br>
'>
'>**Example**
'>
'>なし
'>
'>**See also**
'>
'>なし
'>
Public Sub Clear()
    
    Erase mstrBuf
    ReDim Preserve mstrBuf(0 To mlngInitCount)
    
    mlngCount = 0
    mlngLength = 0

End Sub
'>---
'>#### Append(StringBuilder)
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
'>なし
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
'>Set SB = New StringBuilder
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
'>なし
'>
Function Append(ByVal s As String, Optional ByVal quotation As Boolean = False)

    '配列にセット
    If mlngCount > UBound(mstrBuf) Then
        ReDim Preserve mstrBuf(0 To mlngCount)
    End If
    
    If quotation Then
        mstrBuf(mlngCount) = """" & s & """"
    Else
        mstrBuf(mlngCount) = s
    End If
    
    '文字数をカウント
    mlngLength = mlngLength + Len(mstrBuf(mlngCount))
    
    '要素数＋１
    mlngCount = mlngCount + 1

End Function
'>---
'>#### ToString(StringBuilder)
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
'>Set SB = New StringBuilder
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
'>Set SB = New StringBuilder
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
'>なし
'>
Function ToString(Optional ByVal Delimiter As String = "", Optional ByVal Header As String = "", Optional ByVal Footer As String = "")

    Call resize
    ToString = Header & Join(mstrBuf, Delimiter) & Footer

End Function
'>---
'>#### ToArray(StringBuilder)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToArray()
'>```
'>
'>**Parameters**
'>
'>なし
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
'>Set SB = New StringBuilder
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
'>なし
'>
Function ToArray() As String()
    
    Call resize
    ToArray = mstrBuf()

End Function
'>---
'>#### Length(StringBuilder)
'>
'>**Syntax**
'>
'>```
'>var = expression.Length()
'>```
'>
'>**Parameters**
'>
'>なし
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
'>Set SB = New StringBuilder
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
'>なし
'>
Public Property Get Length() As Long
    Length = mlngLength
End Property
'>---
'>#### Replace(StringBuilder)
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
'>なし
'>
'>**Remarks**
'>
'>置換メソッド<br>
'>
'>**Example**
'>
'>```
'>Dim SB As StringBuilder
'>Set SB = New StringBuilder
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
'>なし
'>
Public Sub Replace(ByVal strFind As String, ByVal strReplace As String)

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

