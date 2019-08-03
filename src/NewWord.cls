VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewWord"
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
'>### NewWord クラス
'>
'>**Remarks**
'>
'>別プロセスWord起動<br>
'>起動時ではなく最初の参照時にWordを起動するため注意。<br>
'>
'>**Example**
'>
'>```
'>With Constructor(New Using, New Word)
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

Private mWD As Object
'Instance を取得時にオブジェクトを生成する。
Public Property Get GetInstance() As Object
    If mWD Is Nothing Then
        Set mWD = CreateObject("Word.Application")
    End If
    Set GetInstance = mWD
End Property

Private Sub Class_Terminate()
    Call IUsing_Finish
End Sub
Private Sub IUsing_Begin()

End Sub

Private Sub IUsing_Finish()
    
    If Not mWD Is Nothing Then
        mWD.Quit
    End If
    Set mWD = Nothing

End Sub
