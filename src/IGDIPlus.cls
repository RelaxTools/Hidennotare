VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGDIPlus"
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
'>### IGDIPlus インターフェイス
'>
'>**Remarks**
'>
'>GDIPlusの動作を定義したインターフェイス
'>
'>**Definition**
'>
'>```
'>Public Sub DrawFillRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
'>Public Sub DrawFillEllipse(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
'>Public Function DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False) As Boolean
'>Public Sub DrawEllipse(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
'>Public Sub DrawString(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal strFont As String, ByVal strBuf As String, ByVal Size As Long, Optional ByVal style As DrawStringConstants, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
'>Public Function ToPicture() As IPictureDisp
'>Public Sub SaveAsJPEG(ByVal strFile As String, Optional ByVal Quality As Long = 85)
'>Public Sub SaveAsPNG(ByVal strFile As String)
'>```
'>
'-----------------------------------------------------------------------------------------------------


Enum DrawStringConstants
    DrawStringFontStyleRegular = 0
    DrawStringFontStyleBold = 1
    DrawStringFontStyleItalic = 2
    DrawStringFontStyleUnderLine = 4
    DrawStringFontStyleStrikeThrough = 8
End Enum

Option Explicit

Public Sub DrawFillRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
    
End Sub
Public Sub DrawFillEllipse(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
    
End Sub
Public Function DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False) As Boolean
    
End Function
Public Sub DrawEllipse(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)
    
End Sub
Public Sub DrawString(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal strFont As String, ByVal strBuf As String, ByVal size As Long, Optional ByVal style As DrawStringConstants, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)

End Sub
Public Function ToPicture() As IPictureDisp

End Function
Public Sub SaveAsJPEG(ByVal strFile As String, Optional ByVal Quality As Long = 85)

End Sub
Public Sub SaveAsPNG(ByVal strFile As String)

End Sub



