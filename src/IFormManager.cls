VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormManager"
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
'>### IFormManager インターフェイス
'>
'>**Remarks**
'>
'>IFormManager
'>
'>**Definition**
'>
'>```
'>Public Property Get IsCancel() As Boolean
'>Public Property Get IsRunning() As Boolean
'>Public Sub StartRunning()
'>Public Sub StopRunning()
'>Public Sub DispGuidance(ByVal strValue As String)
'>Public Sub StartGauge(ByVal lngValue As Long)
'>Public Sub DisplayGauge(ByVal lngValue As Long)
'>Public Sub doCancel()
'>Public Sub SetHistory(c As ComboBox, ByVal strHistory As String)
'>Public Function CreateHistory(c As ComboBox, Optional ByVal Max As Long = 10) As String
'>```
'>
'>**Sample**
'>
'>* None
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Property Get IsCancel() As Boolean

End Property

Public Property Get IsRunning() As Boolean

End Property

Public Sub StartRunning()

End Sub

Public Sub StopRunning()

End Sub

Public Sub DispGuidance(ByVal strValue As String)

End Sub

Public Sub StartGauge(ByVal lngValue As Long)

End Sub

Public Sub DisplayGauge(ByVal lngValue As Long)
    
End Sub

Public Sub doCancel()

End Sub

Public Sub SetHistory(c As ComboBox, ByVal strHistory As String)

End Sub

Public Function CreateHistory(c As ComboBox, Optional ByVal Max As Long = 10) As String

End Function


