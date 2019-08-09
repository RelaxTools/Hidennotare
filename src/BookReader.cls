VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'>### BookReader クラス
'>
'>**Remarks**
'>
'>ブック参照／編集クラス<br>
'>With で使用しブックのオープン。End With でクローズ。<br>
'>別プロセスのExcelで使用する場合はApplicationオブジェクトを渡す。<br>
'>With Constructor(New BookReader, ファイル名, 読み取り専用, Applicationオブジェクト)<br>
'>
'>**Dependent**
'>
'>* IConstructor
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor

Public Enum BookReaderConstants
    None = 0
    SetA1
End Enum

Private XL As Excel.Application
Private WB As Workbook

Private mOpen As Boolean
Private mSheetPos As Long
Private mEof As Boolean
Private Sub Class_Initialize()
    mOpen = False
End Sub
Private Sub Class_Terminate()
    CloseFile
End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret Constructor(New BookReader,
'>                    FileName,
'>                    [Readonly = True],
'>                    [Application = Excel.Application])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を設定|
'>|Readonly|省略可|Boolean|読み取り専用の設定。デフォルトTrue|
'>|Application|省略可|Object|Applicationオブジェクト。デフォルトApplication|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|BookReader|新しいインスタンス(BookReader)|
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>With で生成した場合、End With でブックの Close をします。<br>
'>
'>**Example**
'>
'>```
'>Dim strFile As String
'>Dim WS As Worksheet
'>
'>strFile = "C:\Test.xlsx"
'>
'>With Constructor(New BookReader, strFile)
'>
'>    For Each WS In .VisibleSheets
'>
'>        Debug.Print WS.Name
'>
'>    Next
'>
'>End With
'>```
'>
'>**See also**
'>
'>* None
'>
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Dim strBook As String
    Dim blnReadOnly As Boolean

    Select Case Args.Count
        Case 1
            If TypeOf Args(1) Is Workbook Then
                Set WB = Args(1)
            Else
                strBook = Args(1)       'Book Name
                blnReadOnly = True      'Read Only
                Set XL = Application    'Excel Application
                Call OpenFile(strBook, blnReadOnly)
            End If
        Case 2
            strBook = Args(1)       'Book Name
            blnReadOnly = Args(2)   'Read Only
            Set XL = Application    'Excel Application
            Call OpenFile(strBook, blnReadOnly)
        Case 3
            strBook = Args(1)       'Book Name
            blnReadOnly = Args(2)   'Read Only
            Set XL = Args(3)        'Excel Application
            Call OpenFile(strBook, blnReadOnly)
        Case Else
            Exit Function
    End Select
    
    Set IConstructor_Instancing = Me

End Function
'>---
'>#### VisibleSheets
'>
'>**Syntax**
'>
'>```
'>Set col = object.VisibleSheets()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|col|Collection|表示シートのコレクションを返却|
'>
'>**Remarks**
'>
'>表示シートCollection<br>
'>表示シートのみのコレクションを返却する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Property Get VisibleSheets() As Collection

    Dim WS As Worksheet
    Dim col As Collection
    Set col = New Collection
    
    For Each WS In WB.Worksheets
        If WS.visible = xlSheetVisible Then
            col.Add WS, WS.Name
        End If
    Next
    
    Set VisibleSheets = col

End Property
'-------------------------------------------------
'  ブック Open
'-------------------------------------------------
Private Sub OpenFile(ByVal strBook As String, ByVal blnReadOnly As Boolean)

    If Not FileIO.FileExists(strBook) Then
        Err.Raise 512 + 1, TypeName(Me) & ".OpenFile", "not exists file. (" & strBook & ")"
    End If

    Dim objWB As Workbook
    For Each objWB In XL.Workbooks
        If LCase(objWB.FullName) = LCase(strBook) Then
            Err.Raise 512 + 1, TypeName(Me) & ".OpenFile", "already opened excel file. (" & strBook & ")"
        End If
    Next
    
    Set WB = XL.Workbooks.Open(FileName:=strBook, ReadOnly:=blnReadOnly, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, notify:=False, local:=True)
    If XL.visible Then
        WB.Application.WindowState = xlMinimized
        DoEvents
    End If
    
    mOpen = True

End Sub
'-------------------------------------------------
'  ブック Close
'-------------------------------------------------
Private Sub CloseFile()
    
    If mOpen Then
        On Error Resume Next
        WB.Close SaveChanges:=False
        Set WB = Nothing
        mOpen = False
    End If

End Sub
'>---
'>#### Save
'>
'>**Syntax**
'>
'>```
'>object.Save([opt = BookReaderConstants.None])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|opt|必須|String|保存時にカーソルをA1にする。:BookReaderConstants.SetA1に設定<br>なし:BookReaderConstants.None<br>デフォルト:なし|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>現在のブックを保存する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub Save(Optional ByVal opt As BookReaderConstants = BookReaderConstants.None)
    
    If Not WB.ReadOnly Then
        'お作法
        If opt = BookReaderConstants.SetA1 Then
            setAllA1
        End If
        
        XL.DisplayAlerts = False
        WB.Save
        XL.DisplayAlerts = True
    End If

End Sub
'--------------------------------------------------------------
'　すべてのシートの選択位置をＡ１にセット
'--------------------------------------------------------------
Private Sub setAllA1()

    On Error Resume Next
    
    Dim WS As Worksheet
    Dim WD As Window
    
    For Each WS In WB.Worksheets
        
        If WS.visible = xlSheetVisible Then
            
            WS.Activate
            WS.Range("A1").Activate
            WB.Windows(1).ScrollRow = 1
            WB.Windows(1).ScrollColumn = 1
            WB.Windows(1).Zoom = 100
        
        End If
    Next

    For Each WS In WB.Worksheets
        If WS.visible = xlSheetVisible Then
            WS.Select
            Exit For
        End If
    Next
    
End Sub

