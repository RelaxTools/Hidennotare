VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2
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
'>### Parser クラス
'>
'>**Remarks**
'>
'>JSON文字列からオブジェクトをパースするクラスです。<br>
'>ArrayListやDictionaryからJSONに変換するには各クラスのToStringメソッドで。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* Message
'>* StringBuilder
'>* ArrayList
'>* Dictionary
'>* IList
'>* IDictionary
'>* INewInstance
'>* ICompatibleProperty
'>* StringHelper
'>* Convert
'>* ICursor
'>* CharCursor
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### ParseJSON(Parser)
'>
'>**Syntax**
'>
'>```
'>Set ret = Parser.ParseJSON(strJSON, [obj = Nothing])
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
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|変換後のオブジェクト|
'>
'>**Remarks**
'>
'>JSON文字列からDictionary(OrderedDictionary/SortedDictionary) または ArrayList を復元する機能。<br>
'>ICompatibleProperty(I/F)とINewInstance(I/F)に対応すれば、通常のクラスにも復元可能。
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
'>* ArrayList
'>* Dictionary
'>* OrderdDictionary
'>* SortedDictionary
'>
Public Function ParseJSON(ByVal strJSON As String, Optional ByRef obj As Object = Nothing) As Object

    Dim PL As PairLoggerInfo: Set PL = PairLoggerInfo.CreateObject(TypeName(Me) & ".ParseJSON")
    
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
    Dim SB As IStringBuilder
    
    If Not Core.StartsWith(strJSON, "{") Or Not Core.EndsWith(strJSON, "}") Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseObject", "JSONFormatError"
    End If
    
    'JSONのネストレベル
    Level = Level + 1
    
    Select Case True
        
        'Default
        Case obj Is Nothing
            Set objJObj = CoreEx.GetNewInstance(New Dictionary)
        
        'ネストレベル指定
        Case TypeOf obj Is Collection
            Set objJObj = CoreEx.GetNewInstance(obj(Level))
        
        '直接指定
        Case Else
            Set objJObj = CoreEx.GetNewInstance(obj)
    
    End Select
    
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    Dim IC As ICursor
    Set IC = CharCursor.CreateObject(strJSON)
    
    Do Until IC.Eof
        
        propertyName = ""
        Set SB = StringBuilder.CreateObject()
        lineBreak = False
        
        '終了または行の終わり
        Do Until IC.Eof Or lineBreak
        
            Select Case IC.Item
                Case """"
                    
                    If Len(propertyName) > 0 Then
                        
                        CoreEx.CastICompatibleProperty(objJObj).SetProperty propertyName, ParseString(StringExtraction(IC))
                        propertyName = ""
                    
                    Else
                        propertyName = ParseString(StringExtraction(IC))
                    End If
                    
                Case ":"
                    
                    If Len(propertyName) = 0 Then
                        Err.Raise 512 + 2, TypeName(Me) & ".ParseObject", "JSONFormatError"
                        Exit Function
                    End If
                    
                Case "["
                    
                    '下位レベルのパース（Array)
                    CoreEx.CastICompatibleProperty(objJObj).SetProperty propertyName, ParseArray(ElementExtraction(IC), obj, Level)
                    propertyName = ""
                    
                Case "{"
                    
                    '下位レベルのパース（Object)
                    CoreEx.CastICompatibleProperty(objJObj).SetProperty propertyName, ParseObject(ElementExtraction(IC), obj, Level)
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
                Err.Raise 512 + 3, TypeName(Me) & ".ParseObject", "JSONFormatError"
                Exit Function
            End If
            
            CoreEx.CastICompatibleProperty(objJObj).SetProperty propertyName, ParseValue(SB.ToString)
            
        End If
    
    Loop
    
    Set ParseObject = objJObj

End Function
'-------------------------------------------------
' Arrayのパース
'-------------------------------------------------
Private Function ParseArray(strJSON As String, obj As Object, ByVal Level As Long) As IList
    
    Dim objAry As IList
    Dim lineBreak As Boolean
    Dim SB As IStringBuilder
    
    If Not Core.StartsWith(strJSON, "[") Or Not Core.EndsWith(strJSON, "]") Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseArray", "JSONFormatError"
        Exit Function
    End If
    
    Level = Level + 1
    
    Select Case True
        
        Case obj Is Nothing
            Set objAry = CoreEx.GetNewInstance(New ArrayList)
        
        Case TypeOf obj Is Collection
            Set objAry = CoreEx.GetNewInstance(obj(Level))
        
        Case Else
            Set objAry = CoreEx.GetNewInstance(obj)
    
    End Select
    
    
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    Dim IC As ICursor
    Set IC = CharCursor.CreateObject(strJSON)
    
    Do Until IC.Eof
    
        Set SB = StringBuilder.CreateObject()
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
    
    Dim Buffer As String
    Dim doubleQuat As Long
    
    Dim SB As IStringBuilder
    Set SB = StringBuilder.CreateObject
    
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
        
    Err.raie 512 + 1, TypeName(Me) & "StringExtraction", "JSONFormatError"
End Function
'-------------------------------------------------
' 角カッコまたは波カッコの範囲取得
'-------------------------------------------------
Private Function ElementExtraction(ByRef IC As ICursor) As String
    
    Dim squareBrackets As Long
    Dim curlyBraces As Long
    
    Dim sw As Boolean
    Dim SB As IStringBuilder
    
    sw = False
    Set SB = StringBuilder.CreateObject()
    
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
    
    Err.Raise 512 + 1, TypeName(Me) & ".ElementExtraction", "JSONFormatError"
End Function
'---------------------------------------------------------------------
' 文字列のパース
'---------------------------------------------------------------------
Private Function ParseString(strJSON As String) As String
    
    If Not Core.StartsWith(strJSON, """") Or Not Core.EndsWith(strJSON, """") Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseString", "JSONFormatError"
        Exit Function
    End If
    
    '前後のダブルコーテーションを削除
    strJSON = Mid$(strJSON, 2, Len(strJSON) - 2)
    
    '文字列のUnEscape
    ParseString = Core.Unescape(strJSON)

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


'>---
'>#### ParseFromListbox(Parser)
'>
'>**Syntax**
'>
'>```
'>Set ret = Parser.ParseFromListbox(objList,[obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Object|Parseする ListBox or ComboBox を指定|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|新しいインスタンス(ArrayList)|
'>
'>
'>**Remarks**
'>
'>ListBox/ComboBox パース<br>
'>ListBox/ComboBoxからArrayList(レコードはDictionary)を生成する。<br>
'>Tagにレコードを生成するためのフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    'これはコードではなくプロパティウィンドウで設定しても良い
'>    lstFavorite.Tag = "No,Name,Address"
'>
'>    Dim lst As IList
'>    Set lst = Parser.ParseFromListbox(lstFavorite, New FavoriteBean)
'>
'>```
'>
'>**See also**
'>
'>* Parser.CopyToListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListbox(objList As Object, Optional obj As Object = Nothing) As IList

    Dim ary As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim i As Long
    Dim j As Long
    
    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseFromListbox", "Item name is not set in tag"
    End If
    
    '配列に変換
    Fields = Split(objList.Tag, ",")
    
    '引数によりインスタンス化するオブジェクトを判定
    Select Case True
        Case obj Is Nothing
            Set ary = New ArrayList
        Case Else
            If TypeOf obj Is Collection Then
                Set ary = CoreEx.GetNewInstance(obj(1))
            Else
                Set ary = New ArrayList
            End If
    End Select
        
    For i = 0 To objList.ListCount - 1
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = New Dictionary
            Case Else
                If TypeOf obj Is Collection Then
                    Set objRec = CoreEx.GetNewInstance(obj(2))
                Else
                    Set objRec = CoreEx.GetNewInstance(obj)
                End If
        End Select
        
        For j = 0 To objList.ColumnCount - 1
        
            '設定されているフィールドのみコピーする。
            If Fields(j) <> "" Then
                CoreEx.CastICompatibleProperty(objRec).SetProperty Fields(j), objList.list(i, j)
            End If
        Next
            
        ary.Add objRec
    
    Next

    Set ParseFromListbox = ary

End Function
'>---
'>#### ParseFromListObject(Parser)
'>
'>**Syntax**
'>
'>```
'>Set ret = Parser.ParseFromListObject(objList,[obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Listobject|Parseする Listobject を指定|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|新しいインスタンス(ArrayList)|
'>
'>**Remarks**
'>
'>ListObject パース<br>
'>ListobjectからArrayList(レコードはDictionary)を生成する。<br>
'>ListObject(テーブル)の列ヘッダの情報が必須<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>Sub ArrayList_ParseFromListObject_Sample()
'>
'>    Dim lst As IList
'>    Dim dic As IDictionary
'>    Dim Key As Variant
'>
'>    Set lst = Parser.ParseFromListObject(ActiveSheet.ListObjects(1))
'>
'>    For Each dic In lst
'>
'>        For Each Key In dic.Keys
'>
'>            Debug.Print dic.Item(Key)
'>
'>        Next
'>
'>    Next
'>
'>    Dim a As ArrayList
'>
'>    Set a = lst
'>
'>    a.CopyToListObject ActiveSheet.ListObjects(2)
'>
'>End Sub
'>
'>```
'>
'>**See also**
'>
'>* Parser.CopyToListObject
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListObject(LO As ListObject, Optional obj As Object = Nothing) As IList

    Dim lst As IList
    Dim objRec As Object
    Dim i As Long
    Dim j As Long
    
    '引数によりインスタンス化するオブジェクトを判定
    Select Case True
        Case obj Is Nothing
            Set lst = New ArrayList
        Case Else
            If TypeOf obj Is Collection Then
                Set lst = CoreEx.GetNewInstance(obj(1))
            Else
                Set lst = New ArrayList
            End If
    End Select
        
    For i = 1 To LO.DataBodyRange.Rows.Count
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = New Dictionary
            Case Else
                If TypeOf obj Is Collection Then
                    Set objRec = CoreEx.GetNewInstance(obj(2))
                Else
                    Set objRec = CoreEx.GetNewInstance(obj)
                End If
        End Select
        
        For j = 1 To LO.HeaderRowRange.Columns.Count
        
            CoreEx.CastICompatibleProperty(objRec).SetProperty LO.HeaderRowRange(, j).Value, LO.DataBodyRange.Cells(i, j).Value
        Next
    
        lst.Add objRec
    
    Next

    Set ParseFromListObject = lst

End Function
'--------------------------------------------------------------
' CSVパーサー RFC4180準拠
' コーテーションの有無, デリミタ指定あり
'--------------------------------------------------------------
Public Function ParseCsv(ByVal strBuf As String, Optional ByVal Quatation As Boolean = False, Optional ByVal Delimiter As String = ",", Optional ByVal Header As Boolean = False) As Collection

    Const C_QUAT As String = """"
    Dim IC As ICursor
    Dim sw As Boolean
    Dim blnLineBreak As Boolean
    Dim blnItemBreak As Boolean
    Dim col As Collection
    Dim Row As Collection
    Dim Head As Collection
    Dim first As Boolean
    
    Dim lngQuot As Long
    
    lngQuot = 0
    
    sw = False
    first = True
    
    If Header Then
        Set Head = New Collection
    End If
    
    Set Row = New Collection
    
    Set IC = CharCursor.CreateObject(strBuf)
    Do Until IC.Eof
        
        '初期化
        Set col = New Collection
        blnLineBreak = False

        Do Until IC.Eof Or blnLineBreak
        
            '初期化
            Dim SB As IStringBuilder
            Set SB = StringBuilder.CreateObject()
            blnItemBreak = False
            
            Do Until IC.Eof Or blnLineBreak Or blnItemBreak

                Select Case IC.Item
                    Case C_QUAT

                        'コーテーションありの場合
                        If Quatation Then

                            lngQuot = lngQuot + 1

                            If sw Then

                                '次の文字がQuatation
                                If IC.Item(1) = C_QUAT Then
                                    lngQuot = lngQuot + 1
                                    SB.Append C_QUAT
                                    IC.MoveNext
                                Else
                                    sw = False
                                End If

                            Else
                                sw = True
                            End If
                        End If

                    Case vbCr

                        If sw Then
                            SB.Append IC.Item
                        End If

                    Case vbLf

                        If sw Then
                            SB.Append IC.Item
                        Else
                            blnLineBreak = True
                        End If

                    Case Delimiter

                        If sw Then
                            SB.Append IC.Item
                        Else
                            blnItemBreak = True
                        End If

                    Case Else

                        SB.Append IC.Item

                End Select
            
                IC.MoveNext
            Loop
            
            '列追加処理
            If Header And Not first Then
                col.Add SB.ToString, Head(col.Count + 1)
            Else
                col.Add SB.ToString
            End If
            
        Loop
        
        If Header And first Then
            Set Head = col
        Else
            '行追加処理
            Row.Add col
        End If
        first = False
        
    Loop
    
    'ダブルコーテーションが偶数ではない場合
    If lngQuot Mod 2 <> 0 Then
        Err.Raise 512 + 1, TypeName(Me) & ".CsvParser", "Invalid Format(Quotation)"
    End If
    
    Set ParseCsv = Row
    
End Function

