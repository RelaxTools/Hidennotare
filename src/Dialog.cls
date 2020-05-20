VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dialog"
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
'>### Dialog クラス
'>
'>**Remarks**
'>
'>ダイアログ表示クラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
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
'>|2020-05-20|Ver.1.0.1|ShowFileDialogでフィルタを省略時エラーになるのを修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### ShowColorDialog(Dialog)
'>
'>**Syntax**
'>
'>```
'>ret = Dialog.ShowColorDialog(color)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|color|必須|Long|ダイアログ選択時、色を返却する変数|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|ダイアログで選択された場合Trueを返す。キャンセルした場合はFalse|
'>
'>**Remarks**
'>
'>カラーダイアログ選択<br>
'>ダイアログを表示して選択された色を返却する<br>
'>
'>**Example**
'>
'>```
'>Dim color As long
'>
'>If Dialog.ShowColorDialog(color) Then
'>    lblName.BackColor = color
'>End IF
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function ShowColorDialog(lngDefColor As Long) As Boolean

    Dim lngBackColor As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim strColor As String
    
    ShowColorDialog = False
    
    strColor = Right$("000000" & Hex(lngDefColor), 6)
    lngRed = CLng("&H" & Mid$(strColor, 5, 2))
    lngGreen = CLng("&H" & Mid$(strColor, 3, 2))
    lngBlue = CLng("&H" & Mid$(strColor, 1, 2))
    
    If ActiveWorkbook Is Nothing Then
        Exit Function
    End If
    
    lngBackColor = ActiveWorkbook.Colors(1)
    If Application.Dialogs(xlDialogEditColor).Show(1, lngRed, lngGreen, lngBlue) Then
        lngDefColor = ActiveWorkbook.Colors(1)
        ActiveWorkbook.Colors(1) = lngBackColor
        ShowColorDialog = True
    End If

End Function
'>---
'>#### ShowFolderDialog(Dialog)
'>
'>**Syntax**
'>
'>```
'>ret = Dialog.ShowFolderDialog(folder)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|folder|必須|String|ダイアログ選択時、フォルダ名を返却する変数|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|ダイアログで選択された場合Trueを返す。キャンセルした場合はFalse|
'>
'>**Remarks**
'>
'>フォルダ選択<br>
'>ダイアログを表示して選択されたフォルダ名を返却する<br>
'>
'>**Example**
'>
'>```
'>Dim folder As String
'>
'>'初期表示時のフォルダも指定可能
'>folder = "C:\test"
'>
'>If Dialog.ShowFolderDialog(folder) Then
'>    lblName.text = folder
'>End IF
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function ShowFolderDialog(ByRef strFile As String) As Boolean
    
    ShowFolderDialog = False
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If strFile <> "" Then
            .InitialFileName = FileIO.GetParentFolderName(strFile)
        End If
       
       'タイトルバーの表示文字列の設定
        .Title = "フォルダ選択"
       
       'ボタンの表示文字列の設定
        .ButtonName = "選択"
        
        If .Show Then
            ShowFolderDialog = True
            strFile = .SelectedItems(1)
        End If
    End With

End Function
'>---
'>#### ShowFileDialog(Dialog)
'>
'>**Syntax**
'>
'>```
'>ret = Dialog.ShowFileDialog(file, [filter])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|file|必須|String|ダイアログ選択時、ファイル名を返却する変数|
'>|file|任意|Variant型|ファイルをフィルタしたい場合、バーレンを区切りとしてファイルの種類とワイルドカードを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|ダイアログで選択された場合Trueを返す。キャンセルした場合はFalse|
'>
'>**Remarks**
'>
'>ファイル選択<br>
'>ダイアログを表示して選択されたファイル名を返却する<br>
'>
'>**Example**
'>
'>```
'>Dim strFile As String
'>Dim strFilter As Variant
'>
'>'初期表示時のファイルも指定可能
'>strFile = "C:\test.csv"
'>strFilter = Array("Excel|*.xlsx" ,"Word|*.docx")
'>
'>If Dialog.ShowFileDialog(strFile, strFilter) Then
'>    lblName.text = strFile
'>End IF
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function ShowFileDialog(ByRef strFile As String, Optional ByVal varFilter As Variant) As Boolean
    
    Dim v As Variant
    
    ShowFileDialog = False
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        If strFile <> "" Then
            .InitialFileName = strFile
        End If
        
        .Filters.Clear
        
        If IsMissing(varFilter) Then
            .Filters.Add "すべてのファイル", "*.*"
        Else
            If Not IsArray(varFilter) Then
                varFilter = Array(varFilter)
            End If
            
            For Each v In varFilter
                If InStr(v, "|") > 0 Then
                
                    Dim s As Variant
                    
                    s = Split(v, "|")
                    .Filters.Add CStr(s(0)), CStr(s(1))
                Else
                    .Filters.Add "", CStr(v)
                End If
            Next
        End If
       
       '初期選択フィルタの設定
        .FilterIndex = 1
       
       '複数ファイル選択の許可
        .AllowMultiSelect = False
       
       'タイトルバーの表示文字列の設定
        .Title = "ファイル選択"
       
       'ボタンの表示文字列の設定
        .ButtonName = "選択"
        
        If .Show Then
            ShowFileDialog = True
            strFile = .SelectedItems(1)
        End If
    End With

End Function

