VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneTimeSpeedBooster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'>### OneTimeSpeedBooster クラス
'>
'>**Remarks**
'>
'>マクロ高速化<br>
'>マクロの高速化に寄与するプロパティを設定する｡
'>
'>t-hom’s diary<br>
'>VBA マクロの高速化のためのApplication設定をクラスモジュールにまとめる<br>
'>https://thom.hateblo.jp/entry/2017/03/19/113657<br>
'>
'>**Example**
'>
'>```
'>With Using.CreateObject(New OneTimeSpeedBooster)
'>
'>   'メイン処理
'>
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
'>|2019-08-19|Ver.1.0.1|ブックが存在しない場合にCalculationの指定でエラーになる不具合を修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IUsing

Private Sub IUsing_Begin()
    
    'Applicationのプロパティを変更する。
    With Application
        .ScreenUpdating = False
        
        If Not .ActiveWorkbook Is Nothing Then
            .Calculation = xlCalculationManual
        End If
        
        .EnableEvents = False
        .PrintCommunication = False
        .DisplayAlerts = False
        
        .Cursor = xlWait
    End With

End Sub

Private Sub IUsing_Finish()
    
    'Applicationのプロパティを復元する。
    With Application
        
        .ScreenUpdating = True
        If Not .ActiveWorkbook Is Nothing Then
            .Calculation = xlCalculationAutomatic
        End If
        .EnableEvents = True
        .PrintCommunication = True
        .DisplayAlerts = True
        
        .Cursor = xlDefault
        .StatusBar = False
    
    End With

End Sub

