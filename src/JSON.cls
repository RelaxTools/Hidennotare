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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
' このクラスは厳密なJSONではなくVBAでオブジェクトのシリアライズ／デシリアライズに特化
'-----------------------------------------------------------------------------------------------------
' 依存モジュール
'   なし
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成
'-----------------------------------------------------------------------------------------------------
Option Explicit
'---------------------------------------------------------------
' JSONパース
'---------------------------------------------------------------
Public Function ParseJSON(ByVal strJSON As String, Optional ByVal dic As INewInstance, Optional ByVal lst As INewInstance, Optional ByRef custom As IJSONObject = Nothing) As Object

    Dim PL As PairLoggerInfo: Set PL = Constructor(New PairLoggerInfo, TypeName(Me) & ".ParseJSON")
    
    Select Case Left$(strJSON, 1)
        Case """"
            ParseJSON = ParseString(strJSON)
        Case "{"
            Set ParseJSON = ParseObject(strJSON, dic, lst, custom)
        Case "["
            Set ParseJSON = ParseArray(strJSON, dic, lst, custom)
    End Select

End Function
'-------------------------------------------------
' Objectのパース
'-------------------------------------------------
Public Function ParseObject(ByVal strJSON As String, Optional ByVal dic As INewInstance, Optional ByVal lst As INewInstance, Optional ByVal custom As IJSONObject = Nothing) As Object
    
    Dim objDic As IDictionary
    Dim propertyName As String
    Dim lineBreak As Boolean
    Dim SB As StringBuilder
    
    If Not StringHelper.StartsWith(strJSON, "{") Or Not StringHelper.EndsWith(strJSON, "}") Then
        Message.Throw 1, Me, "ParseObject", "JSONFormatError"
    End If
    
    If dic Is Nothing Then
        Set objDic = New Dictionary
    Else
        Set objDic = dic.NewInstance
    End If
    
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
                        
                        objDic.Add propertyName, ParseString(StringExtraction(IC))
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
                    
                    objDic.Add propertyName, ParseArray(ElementExtraction(IC), dic, lst, custom)
                    propertyName = ""
                    
                Case "{"
                    
                    objDic.Add propertyName, ParseObject(ElementExtraction(IC), dic, lst, custom)
                    propertyName = ""
                
                Case " ", vbTab, vbCr, vbLf, vbCrLf
                
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
            
            objDic.Add propertyName, ParseValue(SB.ToString)
            
        End If
    
    Loop
    
    'カスタムが指定されている場合、変換できるか試行する。
    If Not custom Is Nothing Then
        
        '互換性がある場合コピー
        If custom.IsCompatible(objDic) Then
            Set ParseObject = custom.CopyFromJSONObject(objDic, True)
        Else
            Set ParseObject = objDic
        End If
    
    Else
        Set ParseObject = objDic
    End If
End Function
'-------------------------------------------------
' Arrayのパース
'-------------------------------------------------
Public Function ParseArray(strJSON As String, Optional ByVal dic As INewInstance, Optional ByVal lst As INewInstance, Optional ByVal custom As IJSONObject = Nothing) As IList
    
    Dim objAry As IList
    Dim lineBreak As Boolean
    Dim SB As StringBuilder
    
    If Not StringHelper.StartsWith(strJSON, "[") Or Not StringHelper.EndsWith(strJSON, "]") Then
        Message.Throw 1, Me, "ParseArray", "JSONFormatError"
        Exit Function
    End If
    
    If objAry Is Nothing Then
        Set objAry = New ArrayList
    Else
        Set objAry = lst.NewInstance
    End If
    
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
                    
                    objAry.Add ParseArray(ElementExtraction(IC), dic, lst, custom)
                    
                Case "{"
                    
                    objAry.Add ParseObject(ElementExtraction(IC), dic, lst, custom)
                
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
        
        If IC.Item = """" And IC.PreviousItem <> "\" Then
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
        
        If IC.Item = """" And IC.PreviousItem <> "\" Then
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
Public Function ParseString(strJSON As String) As String
    
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
