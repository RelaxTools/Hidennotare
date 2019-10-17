VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrSch"
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
'>### StrSch クラス
'>
'>**Remarks**
'>
'>文字列検索<br>
'>正規表現検索と互換性を持つ検索クラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* MatchBean
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### Test(StrSch)
'>
'>**Syntax**
'>
'>```
'>flg = StrSch.Test(strText,
'>                  strExp,
'>                  [blnIgnoreCase = False],
'>                  [blnMatchByte = False],
'>                  [blnWhole = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|検索パターンを指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>|blnMatchByte|省略可|Boolean|全角と半角を区別しない場合はTrue、デフォルト値はFalse|
'>|blnWhole|省略可|Boolean|完全一致の場合True、部分一致の場合はFalse、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|マッチ:True / アンマッチ:False|
'>
'>**Remarks**
'>
'>文字列の検索<br>
'>指定された文字列を検索し、パターンに一致する文字列が検索されたかどうかを示す<br>
'>ブール (Boolean) 値を返します。
'>
'>**Example**
'>
'>**See also**
'>
'>* RegExp.Test
'>
Function Test(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False, Optional ByVal blnMatchByte As Boolean = False, Optional ByVal blnWhole As Boolean = False) As Boolean

    Dim strCmp1 As String
    Dim strCmp2 As String
    
    strCmp1 = strText
    strCmp2 = strExp

    If blnIgnoreCase Then
        strCmp1 = UCase(strCmp1)
        strCmp2 = UCase(strCmp2)
    End If

    If blnMatchByte Then
        strCmp1 = StrConv(strCmp1, vbNarrow)
        strCmp2 = StrConv(strCmp2, vbNarrow)
    End If

    If blnWhole Then
        Test = (strCmp1 = strCmp2)
    Else
        Test = (InStr(strCmp1, strCmp2) > 0)
    End If

End Function
'>---
'>#### Replace(StrSch)
'>
'>**Syntax**
'>
'>```
'>ret = StrSch.Replace(strText,
'>                     strExp,
'>                     strReplace,
'>                     [blnIgnoreCase = False],
'>                     [blnMatchByte = False],
'>                     [blnWhole = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|検索パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>|blnMatchByte|省略可|Boolean|全角と半角を区別しない場合はTrue、デフォルト値はFalse|
'>|blnWhole|省略可|Boolean|完全一致の場合True、部分一致の場合はFalse、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|置換後の文字列|
'>
'>**Remarks**
'>
'>文字列の置換<br>
'>指定された文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* RegExp.Replace
'>
Function Replace(ByVal strText As String, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False, Optional ByVal blnMatchByte As Boolean = False, Optional ByVal blnWhole As Boolean = False) As String
    
    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long
    Dim strL As String
    Dim strR As String
    Dim strBuf As String

    '検索
    Set col = Me.Execute(strText, strExp, blnIgnoreCase, blnMatchByte, blnWhole)
    
    strBuf = strText
    '置換するとIndexがずれるので逆順に処理
    For i = col.Count To 1 Step -1
    
        Set bean = col(i)
    
        strL = Mid$(strBuf, 1, bean.Index - 1)
        strR = Mid$(strBuf, bean.Index + bean.Length)

        strBuf = strL & strReplace & strR
    
    Next
    
    Replace = strBuf
    
End Function
'>---
'>#### Execute(StrSch)
'>
'>**Syntax**
'>
'>```
'>Set ret = StrSch.Execute(strText,
'>                         strExp,
'>                         [blnIgnoreCase = False],
'>                         [blnMatchByte = False],
'>                         [blnWhole = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|パターンを指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>|blnMatchByte|省略可|Boolean|全角と半角を区別しない場合はTrue、デフォルト値はFalse|
'>|blnWhole|省略可|Boolean|完全一致の場合True、部分一致の場合はFalse、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Collection<MatchBean>|Collection<MatchBean>を返却|
'>
'>**Remarks**
'>
'>文字列の検索<br>
'>指定された文字列を検索し、パターンに一致する文字列が検索されたかどうかを示す<br>
'>Collectionを返します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StdSch.Execute
'>
Public Function Execute(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False, Optional ByVal blnMatchByte As Boolean = False, Optional ByVal blnWhole As Boolean = False) As Collection

    Dim col As Collection
    Dim lngPos As Long
    Dim r As MatchBean
    Dim strResult As String
    
    Set col = New Collection
    
    Dim strCmp1 As String
    Dim strCmp2 As String
    
    strCmp1 = strText
    strCmp2 = strExp

    If blnIgnoreCase Then
        strCmp1 = UCase(strCmp1)
        strCmp2 = UCase(strCmp2)
    End If

    If blnMatchByte Then
        strCmp1 = StrConv(strCmp1, vbNarrow)
        strCmp2 = StrConv(strCmp2, vbNarrow)
    End If
    
    If blnWhole Then
    
        If strCmp1 = strCmp2 Then
            col.Add MatchBean.CreateObject(1, Len(strCmp1), strCmp1)
        End If
    
    Else
    
        'パターンが空白の場合、全一致とする。
        If strCmp2 = "" Then
            col.Add MatchBean.CreateObject(1, 0, strCmp1)
        Else
    
            lngPos = InStr(strCmp1, strCmp2)
            
            Do Until lngPos = 0
                
                strResult = Mid$(strText, lngPos, Len(strExp))
                
                col.Add MatchBean.CreateObject(lngPos, Len(strResult), strResult)
            
                lngPos = InStr(lngPos + 1, strCmp1, strCmp2)
            Loop
        End If
    End If
    
    Set Execute = col

End Function
'>---
'>#### Check(StrSch)
'>
'>**Syntax**
'>
'>```
'>flg = StrSch.Check(strExp)
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
'>|flg|Boolean|常にTrueを返却|
'>
'>**Remarks**
'>
'>検索文字列のチェック<br>
'>正規表現(RegExp.Check)との互換性のためのメソッド<br>
'>ブール (Boolean) 値を返します。
'>
'>**Example**
'>
'>**See also**
'>
'>* RegExp.Test
'>
Function Check(ByVal strExp As String)
    Check = False
End Function
'>---
'>#### ReplaceShape(StrSch)
'>
'>**Syntax**
'>
'>```
'>StrSch.ReplaceShape sh,
'>                    strText,
'>                    strExp,
'>                    strReplace,
'>                    [blnIgnoreCase = False],
'>                    [blnMatchByte = False],
'>                    [blnWhole = False]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|sh|必須|Shape|置換するShepeを指定|
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|検索パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>|blnMatchByte|省略可|Boolean|全角と半角を区別しない場合はTrue、デフォルト値はFalse|
'>|blnWhole|省略可|Boolean|完全一致の場合True、部分一致の場合はFalse、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ShapeのReplace(書式維持Ver)<br>
'>指定された文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* RegExp.ReplaceShape
'>
Sub ReplaceShape(ByRef sh As Shape, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False, Optional ByVal blnMatchByte As Boolean = False, Optional ByVal blnWhole As Boolean = False)

    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long

    With sh.TextFrame2.TextRange

        '検索
        Set col = Me.Execute(.Text, strExp, blnIgnoreCase, blnMatchByte, blnWhole)
        
        '置換するとIndexがずれるので逆順に処理
        For i = col.Count To 1 Step -1
        
            Set bean = col(i)
        
            Call .Characters(bean.Index, bean.Length).Delete
            
            If strReplace <> "" Then
                Call .Characters(bean.Index, 0).InsertAfter(strReplace)
            End If
        
        Next
    
    End With

End Sub
'>---
'>#### ReplaceRange(StrSch)
'>
'>**Syntax**
'>
'>```
'>StrSch.ReplaceRange rng,
'>                    strText,
'>                    strExp,
'>                    strReplace,
'>                    [blnIgnoreCase = False],
'>                    [blnMatchByte = False],
'>                    [blnWhole = False]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|rng|必須|Range|置換するRangeを指定|
'>|strText|必須|String|検索ターゲットの文字列を指定|
'>|strExp|必須|String|検索パターンを指定|
'>|strReplace|必須|String|置換文字列を指定|
'>|blnIgnoreCase|省略可|Boolean|大文字と小文字を区別しない場合はTrue、デフォルト値はFalse|
'>|blnMatchByte|省略可|Boolean|全角と半角を区別しない場合はTrue、デフォルト値はFalse|
'>|blnWhole|省略可|Boolean|完全一致の場合True、部分一致の場合はFalse、デフォルト値はFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>RangeのReplace(書式維持Ver)<br>
'>指定された文字列を検索し、パターンに一致する文字列が指定文字列で置換<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* RegExp.ReplaceRange
'>
Sub ReplaceRange(ByRef rng As Range, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False, Optional ByVal blnMatchByte As Boolean = False, Optional ByVal blnWhole As Boolean = False)

    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long

    With rng

        '検索
        Set col = Me.Execute(.Value, strExp, blnIgnoreCase, blnMatchByte, blnWhole)
        
        '置換するとIndexがずれるので逆順に処理
        For i = col.Count To 1 Step -1
        
            Set bean = col(i)
        
            Call .Characters(bean.Index, bean.Length).Delete
            
            If strReplace <> "" Then
                Call .Characters(bean.Index, 0).Insert(strReplace)
            End If
        
        Next
    
    End With

End Sub
