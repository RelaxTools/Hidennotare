VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Convert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
'>### Convert クラス
'>
'>**Remarks**
'>
'>変換クラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* StringBuilder
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
Private Const MB_PRECOMPOSED As Long = &H1
'>---
'>#### FromUTF8(Convert)
'>
'>**Syntax**
'>
'>```
'>ret = Convert.FromUTF8(bytBuf())
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|bytBuf()|必須|byte()|UTF8からString型に変換|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|String型(UTF16LE)|
'>
'>
'>**Remarks**
'>
'>Byte型配列(UTF8)からString型(UTF16LE)に変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Convert.ToUTF8
'>
Function FromUTF8(ByRef bytBuf() As Byte) As String
    
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

    FromUTF8 = strBuf

End Function
'>---
'>#### ToUTF8(Convert)
'>
'>**Syntax**
'>
'>```
'>ret() = Convert.ToUTF8(strBuf)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|String|UTF8からString型に変換|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Byte()|byte配列(UTF8)|
'>
'>**Remarks**
'>
'>String型(UTF16LE)をByte型配列(UTF8)に変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Convert.FromUTF8
'>
Function ToUTF8(ByVal strBuf As String) As Byte()
    
    If Len(strBuf) = 0 Then
        ToUTF8 = ""
        Exit Function
    End If
    
    Dim lngSize As Long
    Dim bytBuf() As Byte
    
    lngSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), 0, 0, 0, 0)
    
    ReDim bytBuf(0 To lngSize - 1)
    
    WideCharToMultiByte CP_UTF8, 0, StrPtr(strBuf), Len(strBuf), VarPtr(bytBuf(0)), lngSize, 0, 0
    
    ToUTF8 = bytBuf

End Function

'>---
'>#### ToHiragana(Convert)
'>
'>**Syntax**
'>
'>```
'>ret = Convert.ToHiragana(strBuf, [flag = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|カタカナからひらがなに変換|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|ひらがなに変換された文字列|
'>
'>**Remarks**
'>
'>UNICODE対応ひらがな→カタカナ変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Convert.ToKatakana
'>
Public Function ToHiragana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToHiragana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H30F6&
    Else
        lngOpt = &H30F4&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = Core.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'カタカナ
            Case &H30A1& To lngOpt, &H30FD&, &H30FE&
            
                lngConv = lngBuf - &H60&
                retBuf(i) = Core.LByte(lngConv)
                retBuf(i + 1) = Core.UByte(lngConv)
            
        End Select
    
    Next
    
    ToHiragana = retBuf()

End Function
'>---
'>#### ToKatakana(Convert)
'>
'>**Syntax**
'>
'>```
'>string = Convert.ToKatakana(strBuf, [flag = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|ひらがなからカタカナに変換|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|カタカナに変換された文字列|
'>
'>**Remarks**
'>
'>UNICODE対応ひらがな→カタカナ変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Convert.ToHiragana
'>
Public Function ToKatakana(ByVal strBuf As String, Optional ByVal flag As Boolean = False) As String

    Dim bytBuf() As Byte
    Dim retBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long
    Dim lngConv As Long
    Dim lngOpt As Long
    
    lngLen = 0
    
    If Len(strBuf) = 0 Then
        ToKatakana = ""
        Exit Function
    End If
    
    bytBuf = strBuf
    retBuf = strBuf
    
    If flag Then
        lngOpt = &H3096&
    Else
        lngOpt = &H3094&
    End If
    
    For i = LBound(bytBuf) To UBound(bytBuf) Step 2
    
        lngBuf = Core.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'ひらがな
            Case &H3041& To lngOpt, &H309D&, &H309E&
            
                lngConv = lngBuf + &H60&
                retBuf(i) = Core.LByte(lngConv)
                retBuf(i + 1) = Core.UByte(lngConv)
            
        End Select
    
    Next
    
    ToKatakana = retBuf()

End Function
'--------------------------------------------------------------
'文字化け対応StrConv(vbUnicode, vbFromUnicodeはサポート外)
'--------------------------------------------------------------
Public Function StrConvU(ByVal strSource As String, Conv As VbStrConv) As String

    Dim i As Long
    Dim strBuf As String
    Dim c As String
    Dim strRet As String
    Dim strBefore As String
    Dim strChr As String
    Dim strNext As String

    strRet = ""
    strBuf = ""
    strBefore = ""
    strNext = ""

    For i = 1 To Len(strSource)

        c = Mid$(strSource, i, 1)
        
        If i = Len(strSource) Then
            strNext = ""
        Else
            strNext = Mid$(strSource, i + 1, 1)
        End If

        Select Case c
            '全角の￥
            Case "￥"
                If (Conv And vbNarrow) > 0 Then
                    strChr = "\"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
           
            '半角の\
            Case "\"
                If (Conv And vbWide) > 0 Then
                    strChr = "￥"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            '全角の濁点、半濁点
            Case "゜", "゛"
                If (Conv And vbNarrow) > 0 Then
                    If c = "゜" Then
                        strChr = "ﾟ"
                    Else
                        strChr = "ﾞ"
                    End If
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
                
            '半角の半濁点
            Case "ﾟ"
                '１つ前の文字
                Select Case strBefore
                    Case "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (Conv And vbWide) > 0 Then
                             strChr = "゜"
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
                
            '半角の濁点
            Case "ﾞ"
                '１つ前の文字
                Select Case strBefore
                    Case "ｳ", "ｶ" To "ｺ", "ｻ" To "ｿ", "ﾀ" To "ﾄ", "ﾊ" To "ﾎ"
                        strBuf = strBuf & c
                    Case Else
                        If (Conv And vbWide) > 0 Then
                            strChr = "゛"
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
            'ヴ
            Case "ヴ"
                If (Conv And vbHiragana) > 0 Then
                    Dim b() As Byte
                    ReDim b(0 To 1)
                    b(0) = &H94
                    b(1) = &H30
                    strChr = b
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            'う゛
            Case "う"
                If strNext = "゛" And (Conv And vbKatakana) > 0 Then
                    strChr = "ヴ"
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & strChr
                    strBuf = ""
                    i = i + 1
                Else
                    strBuf = strBuf & c
                End If

            'ヶヵ
            Case "ヶ", "ヵ"
                If (Conv And vbHiragana) > 0 Then
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If

            'その他
            Case Else
                '第二水準等StrConvで文字化けするものを退避
                If Asc(c) = 63 And c <> "?" Then
                    strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                    strBuf = ""
                Else
                    'う”
                    If AscW(c) = &H3094 Then
                        If Conv = vbKatakana Then
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & "ヴ"
                            strBuf = ""
                        Else
                            strRet = strRet & VBA.StrConv(strBuf, Conv) & c
                            strBuf = ""
                        End If
                    Else
                        strBuf = strBuf & c
                    End If
                End If
        End Select
        
        '１個前の文字
        strBefore = c

    Next

    If strBuf <> "" Then
        strRet = strRet & VBA.StrConv(strBuf, Conv)
    End If

    StrConvU = strRet

End Function
