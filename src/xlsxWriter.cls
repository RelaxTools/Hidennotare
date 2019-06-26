VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xlsxWriter"
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
Option Explicit
Implements IWriter
Implements IConstructor

Private mlngRow  As Long
Private mstrFileName As String

Private mWS As Worksheet
Private mWB As Workbook

Private Function IConstructor_Instancing(Args As Collection) As Object

    If Args.Count = 0 Then
        Exit Function
    End If
    
    If TypeOf Args(1) Is Workbook Then
        Set mWB = Args(1)
        mstrFileName = ""
    Else
        mstrFileName = Args(1)
        Set mWB = Workbooks.Add
        mWB.Application.WindowState = xlMinimized
    End If
    
    Set mWS = mWB.Worksheets(1)

    mlngRow = 1
    
    Set IConstructor_Instancing = Me

End Function
Private Sub Class_Terminate()
    
    Call IWriter_Flush
    
    mWB.Close SaveChanges:=False
    Set mWS = Nothing
    Set mWB = Nothing
End Sub

Private Sub IWriter_Append(ByVal Args As Variant)
    
    Dim v As Variant
    Dim lngCol As Long

    lngCol = 1
    
    If Not IsArray(Args) Then
        mWS.Cells(mlngRow, lngCol).Value = Args
    Else
        For Each v In Args
            mWS.Cells(mlngRow, lngCol).Value = v
            lngCol = lngCol + 1
        Next
    End If
    mlngRow = mlngRow + 1

End Sub
Private Sub IWriter_Flush()

    Application.DisplayAlerts = False
    Select Case mstrFileName
        'WB指定の場合には何もしない
        Case ""
        'ファイル名が一緒なら一度保存しているのでSave
        Case UCase(mstrFileName) = UCase(mWB.FullName)
            mWB.Save
        '初回なら名前を付けて保存
        Case Else
            mWB.SaveAs mstrFileName
    End Select
    Application.DisplayAlerts = True

End Sub


