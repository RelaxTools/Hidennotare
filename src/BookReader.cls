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
'  ブック参照／編集クラス
'-----------------------------------------------------------------------------------------------------
'
' With で使用しブックのオープン。End With でクローズ。
' 別プロセスのExcelで使用する場合はApplicationオブジェクトを渡す。
'
'   With Constructor(New BookReader, ファイル名, 読み取り専用, Applicationオブジェクト)
'
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
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - BookName
' Arg2 - ReadOnly    : True/False(Default True)
' Arg3 - Application : Excel Application(Default Application)
'--------------------------------------------------------------
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
'-------------------------------------------------
' 表示されているシートのコレクションを返す。
'-------------------------------------------------
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
        Message.Throw 1, Me, "OpenFile", "Excelファイルが見つかりません。({0})", FileIO.GetFileName(strBook)
    End If

    Dim objWB As Workbook
    For Each objWB In XL.Workbooks
        If LCase(objWB.FullName) = LCase(strBook) Then
            Message.Throw 2, Me, "OpenFile", "すでにExcelファイルが開かれています。({0})", FileIO.GetFileName(strBook)
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
'--------------------------------------------------------------
'　保存
'--------------------------------------------------------------
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

