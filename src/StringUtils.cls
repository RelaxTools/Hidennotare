VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringUtils"
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
'>### StringUtils クラス
'>
'>**Remarks**
'>
'>文字列連結クラス
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
Option Explicit
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
Private Const MB_PRECOMPOSED As Long = &H1
'>---
'>#### FromUTF8(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.FromUTF8(bytBuf())
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
'>* StringUtils.ToUTF8
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
'>#### ToUTF8(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret() = StringUtils.ToUTF8(strBuf)
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
'>* StringUtils.FromUTF8
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
'>#### ToHiragana(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.ToHiragana(strBuf, [flag = False])
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
'>* StringUtils.ToKatakana
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
    
        lngBuf = Me.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'カタカナ
            Case &H30A1& To lngOpt, &H30FD&, &H30FE&
            
                lngConv = lngBuf - &H60&
                retBuf(i) = Me.LByte(lngConv)
                retBuf(i + 1) = Me.UByte(lngConv)
            
        End Select
    
    Next
    
    ToHiragana = retBuf()

End Function
'>---
'>#### ToKatakana(StringUtils)
'>
'>**Syntax**
'>
'>```
'>string = StringUtils.ToKatakana(strBuf, [flag = False])
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
'>* StringUtils.ToHiragana
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
    
        lngBuf = Me.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'ひらがな
            Case &H3041& To lngOpt, &H309D&, &H309E&
            
                lngConv = lngBuf + &H60&
                retBuf(i) = Me.LByte(lngConv)
                retBuf(i + 1) = Me.UByte(lngConv)
            
        End Select
    
    Next
    
    ToKatakana = retBuf()

End Function
'>---
'>#### LenEx(StringUtils)
'>
'>**Syntax**
'>
'>```
'>length = StringUtils.LenEx(strBuf)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|文字数を検査する文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|length|Long|文字数(サロゲートペアを考慮)|
'>
'>**Remarks**
'>
'>サロゲートペア対応Len<br>
'>
'>**Example**
'>
'>* None
'>
Public Function LenEx(ByVal strBuf As String) As Long

    Dim bytBuf() As Byte
    Dim lngBuf As Long
    Dim i As Long
    Dim lngLen As Long

    lngLen = 0

    If Len(strBuf) = 0 Then
        LenEx = 0
        Exit Function
    End If

    bytBuf = strBuf

    For i = LBound(bytBuf) To UBound(bytBuf) Step 2

        lngBuf = Me.LShift(bytBuf(i + 1), 8) + bytBuf(i)

        Select Case lngBuf
            '上位サロゲート
            Case &HD800& To &HDBFF&
                lngLen = lngLen + 1
            '下位サロゲート
            Case &HDC00& To &HDFFF&
                'カウントしない
            Case Else
                lngLen = lngLen + 1
        End Select

    Next

    LenEx = lngLen

End Function
'>---
'>#### TrimEx(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.TrimEx(strBuf)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|文字列|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|左右の空白（全角も含む）を削除した文字列|
'>
'>**Remarks**
'>
'>全角対応トリム<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StringUtils.ToHiragana
'>
Public Function TrimEx(ByVal strBuf As String) As String

    Dim lngLen As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    lngLen = Len(strBuf)

    lngStart = 1
    Do Until lngStart > lngLen
        Select Case Mid$(strBuf, lngStart, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngStart = lngStart + 1
    Loop

    lngEnd = lngLen
    Do Until lngEnd < 1
        Select Case Mid$(strBuf, lngEnd, 1)
            Case Is <= " "
            Case Is = "　"
            Case Else
                Exit Do
        End Select
        lngEnd = lngEnd - 1
    Loop

    If lngEnd > 0 Or lngStart <= lngLen Then
        TrimEx = Mid$(strBuf, lngStart, (lngEnd - lngStart) + 1)
    Else
        TrimEx = ""
    End If

End Function
'>---
'>#### StartsWith(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.StartsWith(string, start)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>string|必須|検査文字列を設定|
'>start|必須|開始文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|指定した文字列で始まる場合True|
'>
'>**Remarks**
'>
'>開始文字列の判定<br>
'>
'>**Example**
'>
'>* None
'>
Public Function StartsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    StartsWith = (strWith = Left$(strBuf, Len(strWith)))

End Function
'>---
'>#### EndsWith(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.StartsWith(string, start)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>string|必須|検査文字列を設定|
'>start|必須|終了文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|指定した文字列で終わる場合True|
'>
'>**Remarks**
'>
'>終了文字列の判定<br>
'>
'>**Example**
'>
'>* None
'>
Public Function EndsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    EndsWith = (strWith = Right$(strBuf, Len(strWith)))

End Function
'>---
'>#### Escape(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.Escape(strBuf)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|String|制御文字をエスケープする文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|エスケープ後の文字列|
'>
'>**Remarks**
'>
'>文字列のエスケープ（正規表現）<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Core.Unescape
'>
Public Function Escape(ByVal v As String) As String

    Dim strChr As String
    Dim SB As IStringBuilder
    Dim i As Long

    Set SB = StringBuilder.NewInstance()

    For i = 1 To Len(v)

        strChr = Mid$(v, i, 1)

        Select Case AscW(strChr)
            Case &H0 To &H7, &HB, &HE To &H1F, &H7F
                SB.Append "\u" & Hex$(AscW(strChr))
            Case &H8
                SB.Append "\b"
            Case &H9
                SB.Append "\t"
            Case &HA
                SB.Append "\n"
            Case &HC
                SB.Append "\f"
            Case &HD
                SB.Append "\r"
            Case &H22
                SB.Append "\"""
            Case &H2F
                SB.Append "\/"
            Case &H5C
                SB.Append "\\"
            Case Else
                SB.Append strChr
        End Select

    Next

    Escape = SB.ToString

End Function
'>---
'>#### Unescape(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.Unescape(strBuf)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>strBuf|必須|String|アンエスケープする文字列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|アンエスケープ後の文字列|
'>
'>**Remarks**
'>
'>文字列のアンエスケープ（正規表現）<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Core.Escape
'>
Public Function Unescape(ByVal v As String) As String

    Dim strChr As String
    Dim strBuf As String
    Dim SB As IStringBuilder
    Dim i As Long

    Set SB = StringBuilder.NewInstance()
    i = 1
    Do Until i > Len(v)

        strChr = Mid$(v, i, 2)

        Select Case strChr
            Case "\u"

                If i + 6 - 1 > Len(v) Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Unescape", "format error"
                End If

                SB.Append ChrW$(val("&H" & Mid$(v, i + 2, 4)))
                i = i + 6
            Case "\b"
                SB.Append Chr$(&H8)
                i = i + 2
            Case "\t"
                SB.Append vbTab
                i = i + 2
            Case "\n"
                SB.Append vbLf
                i = i + 2
            Case "\f"
                SB.Append Chr$(&HC)
                i = i + 2
            Case "\r"
                SB.Append vbCr
                i = i + 2
            Case "\"""
                SB.Append """"
                i = i + 2
            Case "\/"
                SB.Append "/"
                i = i + 2
            Case "\\"
                SB.Append "\"
                i = i + 2
            Case Else
                SB.Append Left$(strChr, 1)
                i = i + 1
        End Select

    Loop

    Unescape = SB.ToString

End Function
'---------------------------------------
' プレースホルダ変換
'---------------------------------------
Public Function PlaceHolder(ByVal strMsg As String, ParamArray p() As Variant) As String

    '翻訳
    strMsg = Me.Unescape(strMsg)

    If UBound(p) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(p)
            strMsg = VBA.Replace(strMsg, "{" & CStr(i) & "}", p(i))
        Next
    End If

    PlaceHolder = strMsg

End Function
'>---
'>#### UByte(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.UByte(value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>value|必須|Long|16bitの文字コード|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|上位8ビットを返却|
'>
'>**Remarks**
'>
'>上位8ビットを返却<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StringUtils.LByte
'>
Public Function UByte(ByVal lngValue As Long) As Long
    UByte = RShift((lngValue And &HFF00&), 8)
End Function
'>---
'>#### LByte(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.LByte(value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>value|必須|Long|16bitの文字コード|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|下位8ビットを返却|
'>
'>**Remarks**
'>
'>下位8ビットを返却<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StringUtils.UByte
'>
Public Function LByte(ByVal lngValue As Long) As Long
    LByte = lngValue And &HFF&
End Function
'>---
'>#### LShift(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.LShift(value, digit)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>value|必須|Long|16～32bitの文字コード|
'>digit|必須|Long|左シフトする桁数|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|シフト結果|
'>
'>**Remarks**
'>
'>左シフト<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StringUtils.RShift
'>
Public Function LShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    LShift = lngValue * (2 ^ lngKeta)
End Function
'>---
'>#### RShift(StringUtils)
'>
'>**Syntax**
'>
'>```
'>ret = StringUtils.RShift(value, digit)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>value|必須|Long|16～32bitの文字コード|
'>digit|必須|Long|右シフトする桁数|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|シフト結果|
'>
'>**Remarks**
'>
'>右シフト<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* StringUtils.LShift
'>
Public Function RShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    RShift = lngValue \ (2 ^ lngKeta)
End Function
'--------------------------------------------------------------
'　文字列のバイト数を求める。漢字２バイト、半角１バイト。
'--------------------------------------------------------------
Public Function AscLen(ByVal var As Variant) As Long

    Dim ascVar As Variant

    ascVar = VBA.StrConv(var, vbFromUnicode)

    AscLen = LenB(ascVar)

End Function
'----------------------------------------------------------------------------------
'　文字列の左端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function AscLeft(ByVal var As Variant, ByVal lngSize As Long) As String

    Dim lngLen As Long
    Dim i As Long

    Dim strChr As String
    Dim strResult As String

    lngLen = Len(var)
    strResult = ""

    For i = 1 To lngLen

        strChr = Mid$(var, i, 1)
        If AscLen(strResult & strChr) > lngSize Then
            Exit For
        End If
        strResult = strResult & strChr

    Next

    AscLeft = strResult

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

