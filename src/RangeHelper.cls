VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeHelper"
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
Option Explicit
'Excelの絶対参照文字列からRangeを生成する。
Function BuildRange(ByVal strBook As String, ByVal strSheet As String, ByVal strAddress) As Range
    
    Dim strPath As String
    
    strPath = lBuildPath(strBook, strSheet, strAddress)
    Set BuildRange = Excel.Range(strPath)
    
End Function
' 選択範囲の定数のセル選択
' 1セル選択の場合に全体の定数セルになってしまうため、Intersectで指定範囲のみにする。
Function SpecialCells(r As Range, c As XlCellType) As Range

    On Error GoTo e

    Set SpecialCells = Me.Intersect(r, r.SpecialCells(c))
    
    Exit Function
e:
    Set SpecialCells = Nothing

End Function
Function SpecialCellsContents(v As Range) As Range

    Dim r As Range
    Dim c As Range
    Dim f As Range

    On Error Resume Next

    Set c = v.SpecialCells(xlCellTypeConstants)
    Set f = v.SpecialCells(xlCellTypeFormulas)

    Set r = Nothing

    If Not c Is Nothing Then
        Set r = c
    End If

    If Not f Is Nothing Then
        If r Is Nothing Then
            Set r = f
        Else
            Set r = v.Application.Union(r, f)
        End If
    End If

    Set SpecialCellsContents = r

End Function
'--------------------------------------------------------------
'　Selection(RANGE)が取得できるかどうかチェックする
'--------------------------------------------------------------
Property Get SelectedRange() As Boolean
    
    On Error GoTo e
    
    SelectedRange = False
    
    Select Case True
        Case ActiveWorkbook Is Nothing
            Exit Property
        Case ActiveCell Is Nothing
            Exit Property
        Case Selection Is Nothing
            Exit Property
        Case TypeOf Selection Is Range
            'OK
        Case Else
            Exit Property
    End Select

    SelectedRange = True

    Exit Property
e:

End Property
'--------------------------------------------------------------
'　Range 同士の和集合を求める(初期値Nothing対応)
'--------------------------------------------------------------
Function Union(r As Range, s As Range) As Range

    If r Is Nothing Then
        Set Union = s
        Exit Function
    End If

    If s Is Nothing Then
        Set Union = r
        Exit Function
    End If

    Set Union = r.Application.Union(r, s)

End Function
'--------------------------------------------------------------
'　Range 同士の積集合を求める
'--------------------------------------------------------------
Function Intersect(r As Range, s As Range) As Range

    If r Is Nothing Then
        Set Intersect = Nothing
        Exit Function
    End If

    If s Is Nothing Then
        Set Intersect = Nothing
        Exit Function
    End If

    Set Intersect = r.Application.Intersect(r, s)

End Function
'--------------------------------------------------------------
'　Range 同士の差集合を求める
'--------------------------------------------------------------
Function Except(r As Range, s As Range) As Range

     Dim PL As PairLogger: Set PL = Constructor(New PairLogger, TypeName(Me) & ".Except")
    '重なる部分があるか？

    Dim n As Range
    
    '共通部分があるか？
    Set n = Me.Intersect(r, s)
    If n Is Nothing Then
        Set Except = r
        GoTo e
    End If
        
    Dim u As Range
    Dim m As Range
    For Each m In r
        '共通部が範囲外なら対象
        If n(1).row <= m.row And n(1).Column <= m.Column And n(n.Count).row >= m.row And n(n.Count).Column >= m.Column Then
        Else
            Set u = Me.Union(u, m)
        End If
    Next
    
    Set Except = u
e:

End Function
'--------------------------------------------------------------
'　A1形式の列IDを番号に変換する。
'--------------------------------------------------------------
Function ConvertToNumber(ByVal strId As String) As Long

    Dim lngColumn As Long
    Dim strChr As String
    Dim i As Long

    For i = 1 To Len(strId)
        
        strChr = UCase(Mid$(strId, Len(strId) - i + 1, 1))
        
        Select Case i
            Case 1
                lngColumn = lngColumn + Asc(strChr) - Asc("@")
            Case 2
                lngColumn = lngColumn + (Asc(strChr) - Asc("@")) * 26
            Case 3
                lngColumn = lngColumn + (Asc(strChr) - Asc("@")) * 26 * 26
        End Select

    Next
    
    ConvertToNumber = lngColumn

End Function
'Private
Private Function lBuildPath(ByVal strBook As String, ByVal strSheet As String, ByVal strAddress) As String
    lBuildPath = "'[" & strBook & "]" & strSheet & "'!" & strAddress
End Function
