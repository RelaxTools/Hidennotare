VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewPowerPoint"
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
'>### NewPowerPoint クラス
'>
'>**Remarks**
'>
'>別プロセスPowerPoint起動<br>
'>Usingクラスと併用してください<br>
'>
'>**Example**
'>
'>```
'>With Using.NewInstance(PowerPoint.NewInstance())
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

Private mPP As Object
'Instance を取得時にオブジェクトを生成する。
Public Function NewInstance() As NewPowerPoint
    
    Set NewInstance = New NewPowerPoint

End Function
Private Sub IUsing_Begin()
    
    Set mPP = VBA.CreateObject("PowerPoint.Application")

End Sub

Private Sub IUsing_Finish()
    
    mPP.Quit
    Set mPP = Nothing

End Sub

Public Function GetInstance() As Object
    
    Set GetInstance = mPP

End Function
