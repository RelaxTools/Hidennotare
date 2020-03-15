VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'>### NewExcel クラス
'>
'>**Remarks**
'>
'>別プロセスExcel起動<br>
'>起動時ではなく最初の参照時にExcelを起動するため注意。<br>
'>
'>**Example**
'>
'>```
'>With Using.NewInstance(New NewExcel)
'>   .Args(1).GetInstance
'>End With
'>```
'>
'>**Dependent**
'>
'>* IUsing
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-05-22|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IUsing

Private mXL As Excel.Application
'Instance を取得時にオブジェクトを生成する。
Public Property Get GetInstance() As Excel.Application
    If mXL Is Nothing Then
        Set mXL = New Excel.Application
        mXL.EnableEvents = False
        mXL.PrintCommunication = False
        mXL.DisplayAlerts = False
    End If
    Set GetInstance = mXL
End Property

Private Sub Class_Terminate()
    Call IUsing_Finish
End Sub

Private Sub IUsing_Begin()

End Sub

Private Sub IUsing_Finish()
    
    If Not mXL Is Nothing Then
        '開いているブックがあれば全部閉じる
        If mXL.Workbooks.Count > 0 Then
            Dim WB As Workbook
            For Each WB In mXL.Workbooks
                WB.Close SaveChanges:=False
            Next
        End If
        mXL.Quit
    End If
    Set mXL = Nothing

End Sub

