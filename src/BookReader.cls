VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookReader"
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
'>### BookReader クラス
'>
'>**Remarks**
'>
'>ブック参照／編集クラス<br>
'>With で使用しブックのオープン。End With でクローズ。<br>
'>別プロセスのExcelで使用する場合はApplicationオブジェクトを渡す。<br>
'>With BookReader.NewInstance(ファイル名, 読み取り専用, Applicationオブジェクト)<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* IBookReader
'>* IBookReaderParam
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにNewInstance以外表示されないように修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IBookReader
Implements IParam

Private XL As Excel.Application
Private wb As Workbook

Private mBookName As String
Private mReadOnly As Boolean

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
'>Set objct = BookReader.NewInstance(FileName,
'>                                   [Readonly = True],
'>                                   [Application = Excel.Application])
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
'>|objct|IBookReader|新しいインスタンス(IBookReader)|
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
'>With BookReader.NewInstance(strFile)
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
Public Function NewInstance(ByVal BookName As String, Optional ByVal ReadOnly As Boolean = True, Optional App As Application = Nothing) As IBookReader
    
    Dim p As IParam
    
    Set p = New BookReader
    
    p.Parameters = Array(BookName, ReadOnly, App)
    
    Set NewInstance = p

End Function
Private Property Let IParam_Parameters(ByVal Args As Variant)

    mBookName = Args(0)
    mReadOnly = Args(1)
    
    If Args(2) Is Nothing Then
        Set XL = Excel.Application
    Else
        Set XL = Args(2)
    End If
    
    OpenFile

End Property
'>---
'>#### VisibleSheets(IBookReader I/F)
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
Private Property Get IBookReader_VisibleSheets() As Collection

    Dim WS As Worksheet
    Dim col As Collection
    Set col = New Collection
    
    For Each WS In wb.Worksheets
        If WS.visible = xlSheetVisible Then
            col.Add WS, WS.Name
        End If
    Next
    
    Set IBookReader_VisibleSheets = col

End Property
'-------------------------------------------------
'  ブック Open
'-------------------------------------------------
Private Sub OpenFile()

    If Not FileIO.FileExists(mBookName) Then
        Err.Raise 512 + 1, TypeName(Me) & ".OpenFile", "not exists file. (" & mBookName & ")"
    End If

    Dim objWB As Workbook
    For Each objWB In XL.Workbooks
        If LCase(objWB.FullName) = LCase(mBookName) Then
            Err.Raise 512 + 1, TypeName(Me) & ".OpenFile", "already opened excel file. (" & mBookName & ")"
        End If
    Next
    
    Set wb = XL.Workbooks.Open(FileName:=mBookName, ReadOnly:=mReadOnly, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, notify:=False, local:=True)
    If XL.visible Then
        wb.Application.WindowState = xlMinimized
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
        wb.Close SaveChanges:=False
        Set wb = Nothing
        mOpen = False
    End If

End Sub
'>---
'>#### Save(IBookReader I/F)
'>
'>**Syntax**
'>
'>```
'>object.Save([opt = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|opt|必須|String|保存時にカーソルをA1にする。:True<br>なし:False<br>デフォルト:なし|
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
Private Sub IBookReader_Save(Optional ByVal opt As Boolean = False)
    
    If Not wb.ReadOnly Then
        'お作法
        If opt Then
            setAllA1
        End If
        
        XL.DisplayAlerts = False
        wb.Save
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
    
    For Each WS In wb.Worksheets
        
        If WS.visible = xlSheetVisible Then
            
            WS.Activate
            WS.Range("A1").Activate
            wb.Windows(1).ScrollRow = 1
            wb.Windows(1).ScrollColumn = 1
            wb.Windows(1).Zoom = 100
        
        End If
    Next

    For Each WS In wb.Worksheets
        If WS.visible = xlSheetVisible Then
            WS.Select
            Exit For
        End If
    Next
    
End Sub


