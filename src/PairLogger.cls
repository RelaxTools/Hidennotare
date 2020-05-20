VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PairLogger"
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
'>### PairLogger クラス
'>
'>**Remarks**
'>
'>開始終了ログ出力(Trace)<br>
'>関数の開始／終了のメッセージを１行で出力するためのクラス<br>
'>Traceレベルで出力する。
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>
'>**Example**
'>
'>```
'>Dim PL As PairLogger
'>Set PL = PairLogger.NewInstance(TypeName(Me) & ".FileSearchEx")
'>```
'>
'>**Dependent**
'>
'>* Logger
'>* IPairLoggerParam
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-05-22|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにNewInstance以外表示されないように修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IPairLoggerParam

Dim mMsg As String
'----------------------------------------------------
' コンストラクタ
'----------------------------------------------------
Public Function NewInstance(Message As String) As PairLogger
    
    Dim PL As IPairLoggerParam
    
    Set PL = New PairLogger
    PL.Message = Message
    
    Set NewInstance = PL

End Function
Private Property Let IPairLoggerParam_Message(RHS As String)
    mMsg = RHS
End Property
'----------------------------------------------------
' デストラクタ
'----------------------------------------------------
Private Sub Class_Terminate()
    Logger.LogFinish LogLevel.Trace, mMsg
End Sub

