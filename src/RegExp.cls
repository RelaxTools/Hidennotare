VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExp"
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
'-----------------------------------------------------------------------------------------------------
' 依存モジュール
'   MatchBean.cls
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 CreateObjectしなくても使える正規表現
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private RE As Object

Private Sub Class_Initialize()
    Set RE = CreateObject("VBScript.RegExp")
End Sub
Private Sub Class_Terminate()
    Set RE = Nothing
End Sub
Public Property Get GetInstance() As Object
    Set GetInstance = RE
End Property
'--------------------------------------------------------------
'  正規表現の検索 True:マッチ/False:アンマッチ
'--------------------------------------------------------------
Function Test(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False) As Boolean

    With RE
        .Pattern = strExp
        .IgnoreCase = blnIgnoreCase
        .Global = False '１つでもマッチすれば終了
        Test = .Test(strText)
    End With

End Function
'--------------------------------------------------------------
'  正規表現の置換
'--------------------------------------------------------------
Function Replace(ByVal strText As String, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False) As String

    strReplace = Convert.Escape(strReplace)
    
    With RE
        .Pattern = strExp
        .IgnoreCase = blnIgnoreCase
        .Global = True '全体を検索
        Replace = .Replace(strText, strReplace)
    End With

End Function
'--------------------------------------------------------------
'  正規表現検索 Collectionを返却
'--------------------------------------------------------------
Function Execute(ByVal strText As String, ByVal strExp As String, Optional ByVal blnIgnoreCase As Boolean = False) As Collection

    Dim col As Collection
    Dim matches As Object
    Dim m As Object
    Dim r As MatchBean
    
    Set col = New Collection

    With RE
        
        .Pattern = strExp
        .IgnoreCase = blnIgnoreCase
        .Global = True  '全体を検索
        
        Set matches = .Execute(strText)
    
    End With
        
    For Each m In matches
    
        'FirstIndexは0からなので+1
        col.Add Constructor(New MatchBean, m.FirstIndex + 1, m.Length, m.Value)
    
    Next
        
    Set Execute = col
    
End Function
'--------------------------------------------------------------
'  正規表現のチェック True:エラー/False:正常
'  正規表現をユーザに入力させる場合に使用
'--------------------------------------------------------------
Function Check(ByVal strExp As String)

    On Error GoTo e
    
    Check = True

    With RE
        .Pattern = strExp
        .IgnoreCase = False
        .Global = False
        .Execute ""
    End With
    
    Check = False
    
    Exit Function
e:

End Function
''--------------------------------------------------------------
'　ShapeのReplace(書式維持Ver)
'--------------------------------------------------------------
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

            Call .Characters(bean.Index, bean.Length).Delete

            If strReplace <> "" Then
                Call .Characters(bean.Index, 0).InsertAfter(strReplace)
            End If

        Next

    End With

End Sub
'--------------------------------------------------------------
'　RangeのReplace(書式維持Ver)
'--------------------------------------------------------------
Sub ReplaceRange(ByRef r As Range, ByVal strExp As String, ByVal strReplace As String, Optional ByVal blnIgnoreCase As Boolean = False)

    Dim col As Collection
    Dim bean As MatchBean
    Dim i As Long

    With r

        '検索
        Set col = Me.Execute(.Value, strExp, blnIgnoreCase)

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

