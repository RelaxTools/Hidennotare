VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'>### 2.2.5 JSON クラス
'>
'>**Remarks**
'>
'>JSON文字列からオブジェクトをパースするクラスです。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* Message クラス
'>* StringBuilder クラス
'>* ArrayList クラス
'>* Dictionary クラス
'>* IList クラス
'>* IDictionary クラス
'>* INewInstance クラス
'>* IDictionaryAccessor クラス
'>* StringHelper クラス
'>* Convert クラス
'>* ICursor クラス
'>* CharCursor クラス
'>* IConstructor クラス
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### ParseJSON(JSON)
'>
'>**Syntax**
'>
'>```
'>Set object = JSON.ParseJSON(strJSON, [obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strJSON|必須|String|JSON文字列を指定|
'>|obj|省略可|object/Collection|オブジェクト（省略可）省略した場合、Object = Dictionary, Array = ArrayList でパースされます。<br>JSON文字列がネストされており、復元するオブジェクトを指定したい場合、Collection にインスタンスを設定したものを指定する。|
'>
'>**Return Value**
'>
'>変換後のオブジェクト
'>
'>**Remarks**
'>
'>JSON文字列からDictionary(OrderedDictionary/SortedDictionary) または ArrayList を復元する機能。<br>
'>IDictionaryAccessor(I/F)とINewInstance(I/F)に対応すれば、通常のクラスにも復元可能。
'>
'>**Example**
'>
'>複数指定する場合、Collectionに詰めて指定する
'>
'>```
'>    Dim col As Collection
'>    Set col = new Collection
'>
'>    col.Add New OrderdDictionary  'ネストの1レベル目に使用されるオブジェクト
'>    col.Add New ArrayList         'ネストの2レベル目に使用されるオブジェクト
'>    col.Add New CustomPerson      'ネストの3レベル目に使用されるオブジェクト
'>
'>    Set dic = JSON.ParseJSON(strJSON, col)
'>```
'>
'>単数指定する場合、Collection を用いず直接指定する
'>
'>```
'>    Set dic = JSON.ParseJSON(strJSON, New CustomPerson)
'>```
'>
'>**See also**
'>
'>* ArrayList クラス
'>* Dictionary クラス
'>* OrderdDictionary クラス
'>* SortedDictionary クラス
'>
Public Function ParseJSON(ByVal strJSON As String, Optional ByRef obj As Object = Nothing) As Object

    Dim PL As PairLoggerInfo: Set PL = Constructor(New PairLoggerInfo, TypeName(Me) & ".ParseJSON")
    
    Dim Level As Long
    
    Level = 0
    
    Select Case Left$(strJSON, 1)
        Case """"
            ParseJSON = ParseString(strJSON)
        Case "{"
            Set ParseJSON = ParseObject(strJSON, obj, Level)
        Case "["
            Set ParseJSON = ParseArray(strJSON, obj, Level)
    End Select

End Function
'-------------------------------------------------
' Objectのパース
'-------------------------------------------------
Private Function ParseObject(ByVal strJSON As String, obj As Object, ByVal Level As Long) As Object
    
    Dim objJObj As Object
    
    Dim propertyName As String
    Dim lineBreak As Boolean
    Dim SB As StringBuilder
    
    If Not StringHelper.StartsWith(strJSON, "{") Or Not StringHelper.EndsWith(strJSON, "}") Then
        Message.Throw 1, Me, "ParseObject", "JSONFormatError"
    End If
    
    'JSONのネストレベル
    Level = Level + 1
    
    Select Case True
        
        'Default
        Case obj Is Nothing
            Set objJObj = GetNewInstance(New Dictionary)
        
        'ネストレベル指定
        Case TypeOf obj Is Collection
            Set objJObj = GetNewInstance(obj(Level))
        
        '直接指定
        Case Else
            Set objJObj = GetNewInstance(obj)
    
    End Select
    
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    Dim IC As ICursor
    Set IC = Constructor(New CharCursor, strJSON)
    
    Do Until IC.Eof
        
        propertyName = ""
        Set SB = New StringBuilder
        lineBreak = False
        
        '終了または行の終わり
        Do Until IC.Eof Or lineBreak
        
            Select Case IC.Item
                Case """"
                    
                    If Len(propertyName) > 0 Then
                        
                        ToIDictionaryAccessor(objJObj).SetProperty propertyName, ParseString(StringExtraction(IC))
                        propertyName = ""
                    
                    Else
                        propertyName = ParseString(StringExtraction(IC))
                    End If
                    
                Case ":"
                    
                    If Len(propertyName) = 0 Then
                        Message.Throw 2, Me, "ParseObject", "JSONFormatError"
                        Exit Function
                    End If
                    
                Case "["
                    
                    '下位レベルのパース（Array)
                    ToIDictionaryAccessor(objJObj).SetProperty propertyName, ParseArray(ElementExtraction(IC), obj, Level)
                    propertyName = ""
                    
                Case "{"
                    
                    '下位レベルのパース（Object)
                    ToIDictionaryAccessor(objJObj).SetProperty propertyName, ParseObject(ElementExtraction(IC), obj, Level)
                    propertyName = ""
                
                Case " ", vbTab, vbCr, vbLf, vbCrLf
                    '無視
                
                Case ","
                    lineBreak = True
                
                Case Else
                    
                    SB.Append IC.Item
            
            End Select
            IC.MoveNext
        Loop
    
        If SB.Length > 0 Then
            
            If Len(propertyName) = 0 Then
                Message.Throw 3, Me, "ParseObject", "JSONFormatError"
                Exit Function
            End If
            
            ToIDictionaryAccessor(objJObj).SetProperty propertyName, ParseValue(SB.ToString)
            
        End If
    
    Loop
    
    Set ParseObject = objJObj

End Function
'-------------------------------------------------
' IDictionaryAccessorへキャスト
'-------------------------------------------------
Private Function ToIDictionaryAccessor(obj As Object) As IDictionaryAccessor
    Set ToIDictionaryAccessor = obj
End Function
'-------------------------------------------------
' インスタンス生成
'-------------------------------------------------
Private Function GetNewInstance(obj As INewInstance) As Object
    Set GetNewInstance = obj.NewInstance
End Function
'-------------------------------------------------
' Arrayのパース
'-------------------------------------------------
Private Function ParseArray(strJSON As String, obj As Object, ByVal Level As Long) As IList
    
    Dim objAry As IList
    Dim lineBreak As Boolean
    Dim SB As StringBuilder
    
    If Not StringHelper.StartsWith(strJSON, "[") Or Not StringHelper.EndsWith(strJSON, "]") Then
        Message.Throw 1, Me, "ParseArray", "JSONFormatError"
        Exit Function
    End If
    
    Level = Level + 1
    
    Select Case True
        
        Case obj Is Nothing
            Set objAry = GetNewInstance(New ArrayList)
        
        Case TypeOf obj Is Collection
            Set objAry = GetNewInstance(obj(Level))
        
        Case Else
            Set objAry = GetNewInstance(obj)
    
    End Select
    
    
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    Dim IC As ICursor
    Set IC = Constructor(New CharCursor, strJSON)
    
    Do Until IC.Eof
    
        Set SB = New StringBuilder
        lineBreak = False
        
        Do Until IC.Eof Or lineBreak
        
            Select Case IC.Item
                Case """"
                    
                    objAry.Add ParseString(StringExtraction(IC))
                    
                Case "["
                    
                    '下位レベルのパース（Array)
                    objAry.Add ParseArray(ElementExtraction(IC), obj, Level)
                    
                Case "{"
                    
                    '下位レベルのパース（Object)
                    objAry.Add ParseObject(ElementExtraction(IC), obj, Level)
                
                Case " ", vbTab, vbCr, vbLf, vbCrLf
                    
                Case ","
                    lineBreak = True
                
                Case Else
                    
                    SB.Append IC.Item
            
            End Select
            
            IC.MoveNext
        
        Loop
    
        If SB.Length > 0 Then
            objAry.Add ParseValue(SB.ToString)
        End If
    
    Loop
    
    Set ParseArray = objAry

End Function
'-------------------------------------------------
' ダブルコーテーションの範囲取得
'-------------------------------------------------
Private Function StringExtraction(ByRef IC As ICursor) As String
    
    Dim buffer As String
    Dim doubleQuat As Long
    
    Dim SB As StringBuilder
    Set SB = New StringBuilder
    
    doubleQuat = 0
    Do Until IC.Eof
    
        SB.Append IC.Item
        
        If IC.Item = """" And IC.Item(-1) <> "\" Then
            doubleQuat = doubleQuat + 1
        End If
        If doubleQuat = 2 Then
            Exit Do
        End If
        
        IC.MoveNext
    
    Loop
    
    If SB.Length > 0 Then
        StringExtraction = SB.ToString
        Exit Function
    End If
        
    Message.Throw 1, Me, "StringExtraction", "JSONFormatError"
End Function
'-------------------------------------------------
' 角カッコまたは波カッコの範囲取得
'-------------------------------------------------
Private Function ElementExtraction(ByRef IC As ICursor) As String
    
    Dim squareBrackets As Long
    Dim curlyBraces As Long
    
    Dim sw As Boolean
    Dim SB As StringBuilder
    
    sw = False
    Set SB = New StringBuilder
    
    Do Until IC.Eof
        
        SB.Append IC.Item
        
        If IC.Item = """" And IC.Item(-1) <> "\" Then
            sw = Not sw
        End If
        
        If Not sw Then
            
            Select Case IC.Item
                Case "["
                    squareBrackets = squareBrackets + 1
                
                Case "]"
                    squareBrackets = squareBrackets - 1
                    
                Case "{"
                    curlyBraces = curlyBraces + 1
                    
                Case "}"
                    curlyBraces = curlyBraces - 1
                
            End Select
        
        End If
        
        If squareBrackets = 0 And curlyBraces = 0 Then
            ElementExtraction = SB.ToString
            Exit Function
        End If
        
        IC.MoveNext
    Loop
    
    Message.Throw 1, Me, "ElementExtraction", "JSONFormatError"
End Function
'---------------------------------------------------------------------
' 文字列のパース
'---------------------------------------------------------------------
Private Function ParseString(strJSON As String) As String
    
    If Not StringHelper.StartsWith(strJSON, """") Or Not StringHelper.EndsWith(strJSON, """") Then
        Message.Throw 1, Me, "ParseString", "JSONFormatError"
        Exit Function
    End If
    
    '前後のダブルコーテーションを削除
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    '文字列のUnEscape
    ParseString = Convert.UnEscape(strJSON)

End Function
'---------------------------------------------------------------------
' 値のパース
'---------------------------------------------------------------------
Private Function ParseValue(ByVal strJSON As String) As Variant

    Select Case True
        Case LCase(strJSON) = "null"
           
           ParseValue = Empty
        
        Case LCase(strJSON) = "true" Or LCase(strJSON) = "false"
           
           ParseValue = CBool(strJSON)
        
        Case IsNumeric(strJSON)
            
            ParseValue = Val(strJSON)
        
        Case Else
            
            ParseValue = strJSON
    
    End Select

End Function


