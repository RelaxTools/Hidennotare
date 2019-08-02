VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PairLoggerInfo"
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
'>### PairLoggerInfo クラス
'>
'>**Remarks**
'>
'>開始終了ログ出力(Info)<br>
'>関数の開始／終了のメッセージを１行で出力するためのクラス<br>
'>Infoレベルで出力する。
'>
'>
'>**Example**
'>
'>```
'>Dim PL As PairLoggerInfo: Set PL = Constructor(New PairLoggerInfo, TypeName(Me) & ".FileSearchEx")
'>```
'>
'>**Dependent**
'>
'>* IConstructor
'>* Logger
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-05-22|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Dim mMsg As String
'----------------------------------------------------
' コンストラクタ
'----------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object
    
    mMsg = Args(1)
    Logger.LogBegin LogLevel.Info, mMsg

    Set IConstructor_Instancing = Me

End Function
'----------------------------------------------------
' デストラクタ
'----------------------------------------------------
Private Sub Class_Terminate()
    Logger.LogFinish LogLevel.Info, mMsg
End Sub

