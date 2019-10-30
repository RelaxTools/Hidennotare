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
'>### BookReader クラス
'>
'>**Remarks**
'>
'>ブック参照／編集クラス<br>
'>With で使用しブックのオープン。End With でクローズ。<br>
'>別プロセスのExcelで使用する場合はApplicationオブジェクトを渡す。<br>
'>With BookReader.CreateObject(ファイル名, 読み取り専用, Applicationオブジェクト)<br>
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
'>|2019-10-28|Ver.1.0.1|クラスのメンバにCreateObject以外表示されないように修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IBookReader
Implements IBookReaderParam

Private XL As Excel.Application
Private WB As Workbook

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
'>Set objct = BookReader.CreateObject(FileName,
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
'>With BookReader.CreateObject(strFile)
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
Public Function CreateObject(ByVal BookName As String, Optional ByVal ReadOnly As Boolean = True, Optional App As Application = Nothing) As IBookReader
    
    Dim BR As IBookReaderParam
    
    Set BR = New BookReader
    
    BR.BookName = BookName
    BR.ReadOnly = ReadOnly
    
    If App Is Nothing Then
        Set BR.Application = Excel.Application
    Else
        Set BR.Application = App
    End If
    
    BR.Init
    
    Set CreateObject = BR

End Function
Private Property Set IBookReaderParam_Application(RHS As Application)
    Set XL = RHS
End Property

Private Property Let IBookReaderParam_BookName(ByVal RHS As String)
    mBookName = RHS
End Property

Private Sub IBookReaderParam_Init()
    OpenFile
End Sub

Private Property Let IBookReaderParam_ReadOnly(ByVal RHS As Boolean)
    mReadOnly = RHS
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
    
    For Each WS In WB.Worksheets
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
    
    Set WB = XL.Workbooks.Open(FileName:=mBookName, ReadOnly:=mReadOnly, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True, notify:=False, local:=True)
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
    
    If Not WB.ReadOnly Then
        'お作法
        If opt Then
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

