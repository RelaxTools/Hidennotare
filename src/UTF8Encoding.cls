VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hecatoncheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/HecatonCheir
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
Option Explicit
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
Private Const MB_PRECOMPOSED As Long = &H1
'------------------------------------------------------------------------------------------------------------------------
' UTF-8 → UTF-16(LE)
'------------------------------------------------------------------------------------------------------------------------
Function GetString(ByRef bytBuf() As Byte) As String
    
    Dim strBuf As String
    Dim lngSize As Long
    Dim lngByte As Long

    If UBound(bytBuf) < 0 Then
        Exit Function
    End If
    
    lngByte = UBound(bytBuf) - LBound(bytBuf) + 1

    lngSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytBuf(0)), lngByte, 0, 0)
    
    strBuf = String$(lngSize, vbNullChar)
    
    lngSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytBuf(0)), lngByte, StrPtr(strBuf), lngSize)

    GetString = strBuf

End Function
'------------------------------------------------------------------------------------------------------------------------
' UTF-16(LE) → UTF-8
'------------------------------------------------------------------------------------------------------------------------
Function GetBytes(ByVal strBuf As String) As Byte()
    
    If Len(strBuf) = 0 Then
        GetBytes = ""
        Exit Function
    End If
    
    Dim lngSize As Long
    Dim bytBuf() As Byte
    
    lngSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), 0, 0, 0, 0)
    
    ReDim bytBuf(0 To lngSize - 1)
    
    WideCharToMultiByte CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), VarPtr(bytBuf(0)), lngSize, 0, 0
    
    GetBytes = bytBuf

End Function
''------------------------------------------------------------------------------------------------------------------------
'' UTF-8 → UTF-16(LE)
''------------------------------------------------------------------------------------------------------------------------
'Public Function GetString(ByRef bytBuf() As Byte) As String
'
'    Dim bytRet() As Byte
'    Dim i As Long
'    Dim lngPos As Long
'    Dim b0 As Long
'    Dim b1 As Long
'    Dim b2 As Long
'    Dim b3 As Long
'
'    GetString = ""
'
'    On Error GoTo e
'
'    i = LBound(bytBuf)
'    ReDim bytRet(0 To (UBound(bytBuf) + 1) * 2)
'    i = 0
'    lngPos = 0
'
'    Do Until i > UBound(bytBuf)
'
'        b0 = bytBuf(i): i = i + 1
'
'        Select Case True
'
''            // UTF-8:   [0xxx xxxx]
''            // Unicode: [0000 0000] [0xxx xxxx]
'            Case (b0 < &H80&)
'                bytRet(lngPos) = b0: lngPos = lngPos + 1
'                bytRet(lngPos) = 0:  lngPos = lngPos + 1
'
''            // UTF-8:   [110y yyyy] [10xx xxxx]
''            // Unicode: [0000 0yyy] [yyxx xxxx]
'            Case ((b0 And &HE0&) = &HC0 And (b0 And &H1E&) <> 0)
'
'                b1 = bytBuf(i): i = i + 1
'                Dim c As Long
'                c = ((Math.LShift(b0, 6)) And &H7C0&) Or (b1 And &H3F&)
'
'                bytRet(lngPos) = LByte(c): lngPos = lngPos + 1
'                bytRet(lngPos) = UByte(c): lngPos = lngPos + 1
'
''            // UTF-8:   [1110 zzzz] [10yy yyyy] [10xx xxxx]
''            // Unicode: [zzzz yyyy] [yyxx xxxx]
'            Case ((b0 And &HF0&) = &HE0&)
'
'                b1 = bytBuf(i): i = i + 1
'                b2 = bytBuf(i): i = i + 1
'
'                c = ((Math.LShift(b0, 12)) And &HF000&) Or ((Math.LShift(b1, 6)) And &HFC0&) Or (b2 And &H3F&)
'
'                bytRet(lngPos) = Math.LByte(c)
'                lngPos = lngPos + 1
'                bytRet(lngPos) = Math.UByte(c)
'                lngPos = lngPos + 1
'
''            // UTF-8:   [1111 0uuu] [10uu zzzz] [10yy yyyy] [10xx xxxx]*
''            // Unicode: [1101 10ww] [wwzz zzyy] (high surrogate)
''            //          [1101 11yy] [yyxx xxxx] (low surrogate)
''            //          * uuuuu = wwww + 1
'            Case ((b0 And &HF8) = &HF0&)
'
'                b1 = bytBuf(i): i = i + 1
'                b2 = bytBuf(i): i = i + 1
'                b3 = bytBuf(i): i = i + 1
'
'                Dim uuuuu As Long
'                Dim wwww As Long
'                Dim zzzz As Long
'                Dim yyyyyy As Long
'                Dim xxxxxx As Long
'                Dim hs As Long
'                Dim ls As Long
'
''                // decode bytes into surrogate characters
'                uuuuu = ((LShift(b0, 2)) And &H1C&) Or ((Math.RShift(b1, 4)) And &H3&)
''                If (uuuuu > &H10) Then
''                    invalidSurrogate(uuuuu);
''                End If
'                wwww = uuuuu - 1
'                zzzz = b1 And &HF&
'                yyyyyy = b2 And &H3F&
'                xxxxxx = b3 And &H3F&
'
'                hs = &HD800& Or ((Math.LShift(wwww, 6)) And &H3C0&) Or (Math.LShift(zzzz, 2)) Or (RShift(yyyyyy, 4))
'                ls = &HDC00& Or ((Math.LShift(yyyyyy, 6)) And &H3C0&) Or xxxxxx
'
'                bytRet(lngPos) = Math.LByte(hs)
'                lngPos = lngPos + 1
'                bytRet(lngPos) = Math.UByte(hs)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = Math.LByte(ls)
'                lngPos = lngPos + 1
'                bytRet(lngPos) = Math.UByte(ls)
'                lngPos = lngPos + 1
'
'        End Select
'
'    Loop
'
'    GetString = LeftB(bytRet, lngPos)
'    Exit Function
'e:
'
'End Function
''------------------------------------------------------------------------------------------------------------------------
'' UTF-16(LE) → UTF-8
''------------------------------------------------------------------------------------------------------------------------
'Public Function getBytes(ByVal strBuf As String) As Byte()
'
'    Dim bytBuf() As Byte
'    Dim lngBuf As Long
'    Dim bytRet() As Byte
'
'    Dim i As Long
'    Dim lngPos As Long
'
'    On Error GoTo e
'
'    If strBuf = "" Then
'        Exit Function
'    End If
'
'    bytBuf = strBuf
'
'    'バッファを最大 1文字×4バイト分確保
'    ReDim bytRet(0 To (Len(strBuf) * 4))
'
'    lngPos = 0
'
'    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
'
'        lngBuf = Math.LShift(bytBuf(i + 1), 8) + bytBuf(i)
'
'        Select Case lngBuf
'
'            Case Is < &H80&
'
'                'UTF-8(ASCII)
'                bytRet(lngPos) = lngBuf
'                lngPos = lngPos + 1
'
'            Case Is < &H800&
'
'                'UTF-8(2バイト)
'                bytRet(lngPos) = &HC0& Or Math.RShift(lngBuf, 6)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
'                lngPos = lngPos + 1
'
'            Case &HD800& To &HDBFF&
'
'                Dim lngHigh As Long
'                Dim lngLow As Long
'
'                lngHigh = lngBuf
'
'                i = i + 2
'                lngLow = Math.LShift(bytBuf(i + 1), 8) + bytBuf(i)
'
'                'サロゲート(UTF-16→Unicode)
'                lngBuf = &H10000 + (lngHigh - &HD800&) * &H400& + (lngLow - &HDC00&)
'
'                'UTF-8(4バイト)
'                bytRet(lngPos) = &HF0& Or Math.RShift(lngBuf, 18)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (Math.RShift(lngBuf, 12) And &H3F&)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (Math.RShift(lngBuf, 6) And &H3F&)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
'                lngPos = lngPos + 1
'
'            Case Else
'
'                'UTF-8(3バイト)
'                bytRet(lngPos) = &HE0& Or Math.RShift(lngBuf, 12)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (Math.RShift(lngBuf, 6) And &H3F&)
'                lngPos = lngPos + 1
'
'                bytRet(lngPos) = &H80& Or (lngBuf And &H3F&)
'                lngPos = lngPos + 1
'
'        End Select
'
'    Next
'
'    getBytes = LeftB(bytRet, lngPos)
'    Exit Function
'e:
'
'End Function

