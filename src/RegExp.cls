VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExp"
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
'>### RegExp クラス
'>
'>**Remarks**
'>
'>CreateObjectしなくても使える正規表現<br>
'>VBScript.RegExpをラップしたクラスです。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* Core
'>* MatchBean
'>* StringBuilder
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private RE As Object

Private Sub Class_Initialize()
    Set RE = VBA.CreateObject("VBScript.RegExp")
End Sub
Private Sub Class_Terminate()
    Set RE = Nothing
End Sub
Public Property Get GetInstance() As Object
    Set GetInstance = RE
End Property
'>---
'>#### Test(RegExp)
'>
'>**Syntax**
'>
'>```
'>boolean = RegExp.Test(strText, strExp, [blnIgnoreCase = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|正規表現パターンを指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|マッチ:True / アンマッチ:False|
'>
'>**Remarks**
'>
'>文字列の正規表現検索<br>
'>指定された正規表現文字列を検索し、パターンに一致する文字列が検索されたかどうかを示す<br>
'>ブール (Boolean) 値を返します。
'>
'>**Example**
'>
'>```
'>If RegExp.Test(strTarget, "[0-9]+") Then
'>    Debug.Print "Number"
'>Else
'>    Debug.Print "Other"
'>End If
'>```
'>
'>**See also**
'>
'>* StrSch.Test
'>
Function Test(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False) As Boolean

    With RE
        .Pattern = strExp
        .ignorecase = blnIgnoreCase
        .Global = False '１つでもマッチすれば終了
        Test = .Test(strText)
    End With

End Function
'>---
'>#### Replace(RegExp)
'>
'>**Syntax**
'>
'>```
'>ret = RegExp.Replace(strText, strExp, strReplace, [blnIgnoreCase = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|正規表現パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|置換後の文字列|
'>
'>**Remarks**
'>
'>文字列の正規表現置換<br>
'>指定された正規表現文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* StrSch.Replace
'>
Function Replace(ByVal strText As String, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False) As String

    strReplace = StringUtils.Escape(strReplace)
    
    With RE
        .Pattern = strExp
        .ignorecase = blnIgnoreCase
        .Global = True '全体を検索
        Replace = .Replace(strText, strReplace)
    End With

End Function
'>---
'>#### Execute(RegExp)
'>
'>**Syntax**
'>
'>```
'>Set ret = RegExp.Execute(strText, strExp, [blnIgnoreCase = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|正規表現パターンを指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Collection<MatchBean>|Collection<MatchBean>を返却|
'>
'>**Remarks**
'>
'>文字列の正規表現検索<br>
'>指定された正規表現文字列を検索し、パターンに一致する文字列が検索されたかどうかを示す<br>
'>Collectionを返します。
'>
'>**Example**
'>
'>```
'>Dim Col As Collection
'>Dim m   As MatchBean
'>Set Col = RegExp.Execute(strTarget, "[0-9]+")
'>
'>For Each m In Col
'>    Debug.Print m.Value
'>    Debug.Print m.Index
'>    Debug.Print m.Length
'>Next
'>```
'>
'>**See also**
'>
'>* StrSch.Execute
'>
Function Execute(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False) As Collection

    Dim col As Collection
    Dim Matches As Object
    Dim m As Object
    Dim r As MatchBean
    
    Set col = New Collection

    With RE
        
        .Pattern = strExp
        .ignorecase = blnIgnoreCase
        .Global = True  '全体を検索
        
        Set Matches = .Execute(strText)
    
    End With
        
    For Each m In Matches
    
        'FirstIndexは0からなので+1
        
        Set r = New MatchBean
        
        r.index = m.FirstIndex + 1
        r.Length = m.Length
        r.Value = m.Value
        
        col.Add r
    
    Next
        
    Set Execute = col
    
End Function
'>---
'>#### Check(RegExp)
'>
'>**Syntax**
'>
'>```
'>flg = RegExp.Check(strExp)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strExp|必須|String|検索パターンを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|エラー:True/正常:False|
'>
'>**Remarks**
'>
'>正規表現のチェック<br>
'>正規表現をユーザに入力させる場合のチェックに使用<br>
'>ブール (Boolean) 値を返します。
'>
'>**Example**
'>
'>**See also**
'>
'>* StrSch.Check
'>
Function Check(ByVal strExp As String)

    On Error GoTo e
    
    Check = True

    With RE
        .Pattern = strExp
        .ignorecase = False
        .Global = False
        .Execute ""
    End With
    
    Check = False
    
    Exit Function
e:

End Function
'>---
'>#### ReplaceShape(RegExp)
'>
'>**Syntax**
'>
'>```
'>RegExp.ReplaceShape sh, strText, strExp, strReplace, [blnIgnoreCase = False]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|sh|必須|Shape|置換を行うShapeを指定|
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|正規表現パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ShapeのReplace(書式維持Ver)<br>
'>指定された正規表現文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* StrSch.ReplaceShape
'>
Sub ReplaceShape(ByRef sh As Shape, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False)

    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long

    With sh.TextFrame2.TextRange

        '検索
        Set col = Me.Execute(.Text, strExp, blnIgnoreCase)

        '置換するとIndexがずれるので逆順に処理
        For i = col.Count To 1 Step -1

            Set bean = col(i)

            Call .Characters(bean.index, bean.Length).Delete

            If strReplace <> "" Then
                Call .Characters(bean.index, 0).InsertAfter(strReplace)
            End If

        Next

    End With

End Sub
'>---
'>#### ReplaceRange(RegExp)
'>
'>**Syntax**
'>
'>```
'>RegExp.ReplaceRange rng, strText, strExp, strReplace, [blnIgnoreCase = False]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|rng|必須|Range|置換を行うRangeを指定|
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|正規表現パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>RangeのReplace(書式維持Ver)<br>
'>指定された正規表現文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* StrSch.ReplaceRange
'>
Sub ReplaceRange(ByRef rng As Range, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False)

    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long

    With rng

        '検索
        Set col = Me.Execute(.Value, strExp, blnIgnoreCase)

        '置換するとIndexがずれるので逆順に処理
        For i = col.Count To 1 Step -1

            Set bean = col(i)

            Call .Characters(bean.index, bean.Length).Delete

            If strReplace <> "" Then
                Call .Characters(bean.index, 0).Insert(strReplace)
            End If

        Next

    End With

End Sub
