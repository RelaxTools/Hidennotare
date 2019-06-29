VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hecatoncheir] v4
'
' Copyright (c) 2009 Yasuhiro Watanabe
' https://github.com/RelaxTools/RelaxTools-Addin
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
Option Explicit
'--------------------------------------------------------------
' 新ブックへのシートのコピー
'--------------------------------------------------------------
Function CopySheetAsNewBook(WS As Worksheet, Optional ByVal Minimized As Boolean = False) As Worksheet

    WS.Copy
    With WS.Application
        Set CopySheetAsNewBook = .Workbooks(.Workbooks.Count).Sheets(1)
    End With

    If Minimized Then
        '最小化
        With CopySheetAsNewBook
            .Parent.Activate
            .Application.WindowState = xlMinimized
        End With
        DoEvents
    End If
    
End Function
'--------------------------------------------------------------
' シートのコピー（シートの最後に追加）
'--------------------------------------------------------------
Function CopySheet(WS As Worksheet) As Worksheet

    WS.Copy After:=WS.Parent.Sheets(WS.Parent.Sheets.Count)
    Set CopySheet = WS.Application.ActiveSheet

End Function
'--------------------------------------------------------------
' シート名が重複しないようにリネーム
'--------------------------------------------------------------
Sub RenameUniqSheet(WS As Worksheet, ByVal strDefName As String)

    Dim blnFind As Boolean
    Dim s As Worksheet
    Dim lngCnt As Long
    Dim strName As String
    
    lngCnt = 1
    strName = strDefName
    
    Do
        blnFind = False
        For Each s In WS.Parent.Worksheets
            If s.Name = strName Then
                blnFind = True
                lngCnt = lngCnt + 1
                strName = strDefName & " (" & lngCnt & ")"
                Exit For
            End If
        Next
    Loop Until blnFind = False
    
    WS.Name = strName
    
End Sub
'--------------------------------------------------------------
' 同一ブックOpenチェック
'--------------------------------------------------------------
Function IsOpenSameBook(ByVal strBook As String) As Boolean
    
    Dim v As Workbook
    
    IsOpenSameBook = False
    
    For Each v In Workbooks
        If LCase(strBook) = LCase(v.Name) Then
            IsOpenSameBook = True
            Exit For
        End If
    Next

End Function
'---------------------------------------------------------------------------------------------------
'　Application.Run / .OnKey / .OnAction / .OnUndo / .OnRepeat / .OnTime の頭にブック名を付ける
'---------------------------------------------------------------------------------------------------
Public Function BuildOnAction(ByVal strMacro As String) As String

    BuildOnAction = "'" & ThisWorkbook.Name & "'!" & strMacro
    
End Function
'--------------------------------------------------------------
'  Excelの絶対参照文字列を生成する。
'--------------------------------------------------------------
Public Function BuildPath(ByVal strBook As String, ByVal strSheet As String, ByVal strAddress) As String
    
    BuildPath = "'[" & strBook & "]" & strSheet & "'!" & strAddress

End Function
'--------------------------------------------------------------
'　Excelファイル判定
'--------------------------------------------------------------
Public Function IsExcelFile(ByVal strFile As String) As Boolean

    Dim varExt As Variant
    Dim i As Long
    IsExcelFile = False
    
    varExt = Array(".XLSX", ".XLSM", ".XLS", ".XLB")

    For i = LBound(varExt) To UBound(varExt)
    
        If Left$(FileIO.GetFileName(strFile), 2) <> "~$" And InStr(UCase(strFile), varExt(i)) > 0 Then
            IsExcelFile = True
            Exit For
        End If
    
    Next

End Function
'--------------------------------------------------------------
'　PowerPointファイル判定
'--------------------------------------------------------------
Public Function IsPowerPointFile(ByVal strFile As String) As Boolean

    Dim varExt As Variant
    Dim i As Long
    IsPowerPointFile = False
    
    varExt = Array(".PPT", ".PPTX")

    For i = LBound(varExt) To UBound(varExt)
    
        If InStr(UCase(strFile), varExt(i)) > 0 Then
            IsPowerPointFile = True
            Exit For
        End If
    
    Next

End Function
'--------------------------------------------------------------
'　Wordファイル判定
'--------------------------------------------------------------
Public Function IsWordFile(ByVal strFile As String) As Boolean

    Dim varExt As Variant
    Dim i As Long
    IsWordFile = False
    
    varExt = Array(".DOC", ".DOCX")

    For i = LBound(varExt) To UBound(varExt)
    
        If InStr(UCase(strFile), varExt(i)) > 0 Then
            IsWordFile = True
            Exit For
        End If
    
    Next

End Function


