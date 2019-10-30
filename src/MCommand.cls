VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------------
'
' [MLangBuilder] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/MFunctionCreater
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
'>### MCommand クラス
'>
'>**Remarks**
'>
'>MCommand<br>
'>Power Query M Language Command Class<br>
'>M言語をVBAで組み立てる際に便利なクラス<br>
'>
'>**Dependent**
'>
'>* StringBuilder
'>* MTable
'>* IMCommand
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IMCommand

'テーブルの別名のデフォルト値
Const TABLE_NAME As String = "Source"

'テーブルの連番
Private mSeq As Long

Private mSb As IStringBuilder

'最後に追加したテーブル
Private mLast As MTable

'別名
Private mAliasName As String
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set obj = MCommand.CreateObject()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|MCommand|必須|MCommand|MCommandを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|obj|MCommand|新しいインスタンス(IMCommand)|
'>
'>**Remarks**
'>
'> MCommand を生成する場合
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
'>---
Public Function CreateObject() As IMCommand

    Set CreateObject = New MCommand
    
    
    
    Call Class_Terminate
End Function
'---------------------------------------
' Initialize
'---------------------------------------
Private Sub Class_Initialize()
    Set mSb = StringBuilder.CreateObject()
    mSeq = 1
    mAliasName = TABLE_NAME
End Sub
'---------------------------------------
' Terminate
'---------------------------------------
Private Sub Class_Terminate()
    Set mSb = Nothing
End Sub
'>---
'>#### AliasName(IMCommand I/F)
'>
'>**Syntax**
'>
'>```
'>expression.AliasName() = value
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MCommand|MCommandインスタンスを指定|
'>|value|String|テーブルの別名|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テーブルの別名<br>
'>
'>**See also**
'>
'>* None
'>
Private Property Let IMCommand_AliasName(v As String)
    mAliasName = v
End Property
'>---
'>#### Append(IMCommand I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Append(table, [aliasName = ""])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MCommand|MCommandインスタンスを指定|
'>|table|必須|MTable|追加するテーブルを指定|
'>|aliasName|省略可|String|テーブルの別名を指定可能。デフォルト""|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テーブル追加メソッド<br>
'>デリミタ、ヘッダ、フッタを指定可能です。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Sub IMCommand_Append(m As MTable, Optional ByVal optName As String = "")

    If optName = "" Then
        If m.Name = "" Then
            m.Name = mAliasName & mSeq
            mSeq = mSeq + 1
        End If
    Else
        m.Name = optName
    End If
     
    mSb.Append m.Name & " = " & m.ToString
    Set mLast = m

End Sub
'>---
'>#### Table(IMCommand I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Table()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MCommand|MCommandインスタンスを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|テーブルの別名|
'>
'>**Remarks**
'>
'>テーブル別名取得プロパティ<br>
'>最後に追加したテーブルを取得<br>
'>
'>**See also**
'>
'>* None
'>
Private Property Get IMCommand_Table() As MTable
    Set IMCommand_Table = mLast
End Property
'>---
'>#### ToString(IMCommand I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToString()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MCommand|MCommandインスタンスを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|MCommandに追加されたテーブルをM言語を生成|
'>
'>**Remarks**
'>
'>M言語変換メソッド<br>
'>追加されたテーブルの順に別名で連結、M言語を生成する<br>
'>
'>**See also**
'>
'>* None
'>
Private Function IMCommand_ToString()

    IMCommand_ToString = "let " & vbCrLf & mSb.ToString(", " & vbCrLf) & " " & vbCrLf & "in " & mLast.Name

End Function

