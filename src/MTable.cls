VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MTable"
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
'>### MTable クラス
'>
'>**Remarks**
'>
'>MTable<br>
'>Power Query M Language Table Class<br>
'>M言語をVBAで組み立てる際に便利なクラス<br>
'>
'* M 言語の Table が作成できるクラス
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Name As String
Private mValue As String
Private mParent As MTable
'---------------------------------------
' コンストラクタ
'---------------------------------------
Public Function CreateObject(Value As String, Optional ByRef Parent As MTable = Nothing) As MTable
    
    Dim m As MTable
    
    Set m = New MTable
    
    m.Value = Value
    Set m.Parent = Parent
    
    Set CreateObject = m

End Function
Public Property Let Value(v As String)
    mValue = v
End Property
Public Property Set Parent(v As MTable)
    Set mParent = v
End Property
'>---
'>#### PromoteHeaders(MFile)
'>
'>**Syntax**
'>
'>```
'>set ret = expression.PromoteHeaders(table, [option])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MTable|MTableインスタンスを指定|
'>|table|必須|MTable|対象のテーブルを指定|
'>|option|省略可|String|オプションを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|MTable|テーブルを返却|
'>
'>**Remarks**
'>
'>Contentsメソッド<br>
'>
'>**See also**
'>
'>* None
'>
Public Function PromoteHeaders(t As MTable, Optional m As Variant) As MTable

    Dim strName As String
    Dim strMFunc As String
    
    If IsMissing(m) Then
        strMFunc = ""
    Else
        If IsObject(m) Then
            strMFunc = ", " & m.ToString()
        Else
            strMFunc = ", " & m
        End If
    End If
    
    If t.Name = "" Then
        strName = t.ToString
    Else
        strName = t.Name
    End If
    
    Set PromoteHeaders = MTable.CreateObject(GetClassName & "PromoteHeaders(" & strName & strMFunc & ")", t)

End Function
'>---
'>#### InsertRows(MTable)
'>
'>**Syntax**
'>
'>```
'>set ret = expression.InsertRows(table, offset, list)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MTable|MTableインスタンスを指定|
'>|table|必須|MTable|対象のテーブルを指定|
'>|offset|必須|Long|行追加を行う位置を指定|
'>|list|必須|Mlist|リストを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|MTable|テーブルを返却|
'>
'>**Remarks**
'>
'>InsertRowsメソッド<br>
'>
'>**See also**
'>
'>* None
'>
Public Function InsertRows(t As MTable, offset As Long, lst As Variant) As MTable

    Dim strName As String
    Dim strMFunc As String
    
    If IsObject(lst) Then
        strMFunc = lst.ToString
    Else
        strMFunc = CStr(lst)
    End If
    
    If t.Name = "" Then
        strName = t.ToString
    Else
        strName = t.Name
    End If
    
    Set InsertRows = MTable.CreateObject(GetClassName & "InsertRows(" & strName & ", " & offset & ", " & strMFunc & ")", t)

End Function
'>---
'>#### TransformColumnTypes(MTable)
'>
'>**Syntax**
'>
'>```
'>set ret = expression.TransformColumnTypes(table, list)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MTable|MTableインスタンスを指定|
'>|table|必須|MTable|対象のテーブルを指定|
'>|list|必須|Mlist|リストを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|MTable|テーブルを返却|
'>
'>**Remarks**
'>
'>TransformColumnTypesメソッド<br>
'>
'>**See also**
'>
'>* None
'>
Public Function TransformColumnTypes(t As MTable, lst As Variant) As MTable

    Dim strName As String
    Dim strMFunc As String
    
    If IsObject(lst) Then
        strMFunc = lst.ToString
    Else
        strMFunc = CStr(lst)
    End If
    
    If t.Name = "" Then
        strName = t.ToString
    Else
        strName = t.Name
    End If
    
    Set TransformColumnTypes = MTable.CreateObject(GetClassName & "TransformColumnTypes(" & strName & ", " & strMFunc & ")", t)

End Function
'>---
'>#### Skip(MTable)
'>
'>**Syntax**
'>
'>```
'>set ret = expression.Skip(table, list)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|MTable|MTableインスタンスを指定|
'>|countOrCondition|必須|Long|スキップする行数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|MTable|テーブルを返却|
'>
'>**Remarks**
'>
'>Skipメソッド<br>
'>
'>**See also**
'>
'>* None
'>
Public Function Skip(t As MTable, countOrCondition As Long) As MTable
    
    Dim strBuf As String
    
    If t.Name = "" Then
        strBuf = t.ToString
    Else
        strBuf = t.Name
    End If
    
    Set Skip = MTable.CreateObject(GetClassName & "Skip(" & strBuf & ", " & countOrCondition & ")", t)
    
End Function
'>---
'>#### ToString(MTable)
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
'>|expression|必須|MTable|MTableインスタンスを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|MTableに追加されたテーブルをM言語を生成|
'>
'>**Remarks**
'>
'>M言語変換メソッド<br>
'>追加されたテーブルのM言語を生成する<br>
'>
'>**See also**
'>
'>* None
'>
Public Function ToString() As String
    ToString = mValue
End Function
'---------------------------------------
' クラス名取得
'---------------------------------------
Private Function GetClassName() As String
    GetClassName = Mid$(TypeName(Me), 2) & "."
End Function

