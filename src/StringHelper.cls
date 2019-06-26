VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hecatoncheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hecatoncheir
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
' 依存モジュール
'-----------------------------------------------------------------------------------------------------
' 依存モジュール
'   Math.bas
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成
'-----------------------------------------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------
'  UNICODE対応ひらがな→カタカナ変換
'--------------------------------------------------------------
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
    
        lngBuf = Math.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'ひらがな
            Case &H3041& To lngOpt, &H309D&, &H309E&
            
                lngConv = lngBuf + &H60&
                retBuf(i) = Math.LByte(lngConv)
                retBuf(i + 1) = Math.UByte(lngConv)
            
        End Select
    
    Next
    
    ToKatakana = retBuf()

End Function
'--------------------------------------------------------------
'  UNICODE対応カタカナ→ひらがな変換
'--------------------------------------------------------------
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
    
        lngBuf = Math.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
        Select Case lngBuf
            'カタカナ
            Case &H30A1& To lngOpt, &H30FD&, &H30FE&
            
                lngConv = lngBuf - &H60&
                retBuf(i) = Math.LByte(lngConv)
                retBuf(i + 1) = Math.UByte(lngConv)
            
        End Select
    
    Next
    
    ToHiragana = retBuf()

End Function
'--------------------------------------------------------------
'  サロゲートペア対応Len
'--------------------------------------------------------------
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
    
        lngBuf = Math.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
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
'--------------------------------------------------------------
'  全角対応トリム
'--------------------------------------------------------------
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
'--------------------------------------------------------------
'文字化け対応StrConv(vbUnicode, vbFromUnicodeはサポート外)
'--------------------------------------------------------------
Public Function StrConvU(ByVal strSource As String, conv As VbStrConv) As String

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
                If (conv And vbNarrow) > 0 Then
                    strChr = "\"
                    strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
           
            '半角の\
            Case "\"
                If (conv And vbWide) > 0 Then
                    strChr = "￥"
                    strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            '全角の濁点、半濁点
            Case "゜", "゛"
                If (conv And vbNarrow) > 0 Then
                    If c = "゜" Then
                        strChr = "ﾟ"
                    Else
                        strChr = "ﾞ"
                    End If
                    strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
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
                        If (conv And vbWide) > 0 Then
                             strChr = "゜"
                            strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
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
                        If (conv And vbWide) > 0 Then
                            strChr = "゛"
                            strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
                            strBuf = ""
                        Else
                            strBuf = strBuf & c
                        End If
                End Select
            'ヴ
            Case "ヴ"
                If (conv And vbHiragana) > 0 Then
                    Dim b() As Byte
                    ReDim b(0 To 1)
                    b(0) = &H94
                    b(1) = &H30
                    strChr = b
                    strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If
            'う゛
            Case "う"
                If strNext = "゛" And (conv And vbKatakana) > 0 Then
                    strChr = "ヴ"
                    strRet = strRet & VBA.StrConv(strBuf, conv) & strChr
                    strBuf = ""
                    i = i + 1
                Else
                    strBuf = strBuf & c
                End If

            'ヶヵ
            Case "ヶ", "ヵ"
                If (conv And vbHiragana) > 0 Then
                    strRet = strRet & VBA.StrConv(strBuf, conv) & c
                    strBuf = ""
                Else
                    strBuf = strBuf & c
                End If

            'その他
            Case Else
                '第二水準等StrConvで文字化けするものを退避
                If Asc(c) = 63 And c <> "?" Then
                    strRet = strRet & VBA.StrConv(strBuf, conv) & c
                    strBuf = ""
                Else
                    'う”
                    If Unicode(c) = &H3094 Then
                        If conv = vbKatakana Then
                            strRet = strRet & VBA.StrConv(strBuf, conv) & "ヴ"
                            strBuf = ""
                        Else
                            strRet = strRet & VBA.StrConv(strBuf, conv) & c
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
        strRet = strRet & VBA.StrConv(strBuf, conv)
    End If

    StrConvU = strRet

End Function
Private Function Unicode(ByVal strBuf As String) As Long
    Dim bytBuf() As Byte
    
    If Len(strBuf) <> 0 Then
        bytBuf = strBuf
        Unicode = CLng(bytBuf(1)) * &H100 + bytBuf(0)
    End If
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
    
        strChr = Mid(var, i, 1)
        If AscLen(strResult & strChr) > lngSize Then
            Exit For
        End If
        strResult = strResult & strChr
    
    Next

    AscLeft = strResult

End Function
'----------------------------------------------------------------------------------
'　文字列の右端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function AscRight(ByVal var As Variant, ByVal lngSize As Long) As String

    Dim ascVar As Variant
    
    ascVar = VBA.StrConv(var, vbFromUnicode)

    AscRight = VBA.StrConv(RightB(ascVar, lngSize), vbUnicode)

End Function
'----------------------------------------------------------------------------------
'　文字列から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
'----------------------------------------------------------------------------------
Public Function AscMid(ByVal var As Variant, ByVal lngPos As Long, Optional ByVal varSize As Variant) As String

    Dim ascVar As Variant
    
    ascVar = VBA.StrConv(var, vbFromUnicode)

    If IsMissing(varSize) Then
        AscMid = VBA.StrConv(MidB(ascVar, lngPos), vbUnicode)
    Else
        AscMid = VBA.StrConv(MidB(ascVar, lngPos, varSize), vbUnicode)
    End If

End Function
'---------------------------------------------------------------------
' エスケープ文字のデコード
'---------------------------------------------------------------------
Public Function EscDecode(ByVal v As String) As String

    Dim strChr As String
    Dim strBuf As String
    Dim sb As StringBuilder
    Dim i As Long
    
    '本当にJSONするならここでUTF8変換
    
    Set sb = New StringBuilder
    i = 1
    Do Until i > Len(v)

        strChr = Mid$(v, i, 2)
        
        Select Case strChr
            Case "\u"
            
                If i + 6 > Len(v) Then
                    Err.Raise 513, , "error"
                End If
            
                sb.Append ChrW$(Val("&H" & Mid$(v, i + 2, 4)))
                i = i + 6
            Case "\b"
                sb.Append Chr$(&H8)
                i = i + 2
            Case "\t"
                sb.Append vbTab
                i = i + 2
            Case "\n"
                sb.Append vbLf
                i = i + 2
            Case "\f"
                sb.Append Chr$(&HC)
                i = i + 2
            Case "\r"
                sb.Append vbCr
                i = i + 2
            Case "\"""
                sb.Append """"
                i = i + 2
            Case "\/"
                sb.Append "/"
                i = i + 2
            Case "\\"
                sb.Append "\"
                i = i + 2
            Case Else
                sb.Append Left(strChr, 1)
                i = i + 1
        End Select

    Loop
    

    EscDecode = sb.ToString

End Function
'--------------------------------------------------------------
'  RGBからCSS表記のカラー文字列に変更
'--------------------------------------------------------------
Function RGBToCSS(ByVal lngColor As Variant) As String

    Dim strBuf As String
    If IsNull(lngColor) Then
        RGBToCSS = "#000000"
    Else
        strBuf = Right$("000000" & Hex$(lngColor), 6)
        RGBToCSS = "#" & Mid$(strBuf, 5, 2) & Mid$(strBuf, 3, 2) & Mid$(strBuf, 1, 2)
    End If
End Function
'--------------------------------------------------------------
'  ＨＴＭＬ文字列のサニタイジングを行う。
'--------------------------------------------------------------
Public Function Sanitizing(ByVal strBuf As String) As String

    Dim strRep As String

    strRep = Replace(strRep, "&", "&amp;")
    strRep = Replace(strBuf, """", "&quot;")
    strRep = Replace(strRep, "<", "&lt;")
    strRep = Replace(strRep, ">", "&gt;")
    strRep = Replace(strRep, "'", "&#39;")
    Sanitizing = strRep

End Function

'----------------------------------------------------------------------------------
'　ローマ数字変換
'----------------------------------------------------------------------------------
Function ToRoman(ByVal lngNum As Long) As String

    Dim lng1000 As Long
    Dim lng500 As Long
    Dim lng100 As Long
    Dim lng50 As Long
    Dim lng10 As Long
    Dim lng5 As Long
    
    Dim strAns As String

    'M
    lng1000 = lngNum \ 1000
    lngNum = lngNum - 1000 * lng1000
    
    'D
    lng500 = lngNum \ 500
    lngNum = lngNum - 500 * lng500
    
    'C
    lng100 = lngNum \ 100
    lngNum = lngNum - 100 * lng100
    
    'L
    lng50 = lngNum \ 50
    lngNum = lngNum - 50 * lng50
    
    'X
    lng10 = lngNum \ 10
    lngNum = lngNum - 10 * lng10
    
    'V
    lng5 = lngNum \ 5
    lngNum = lngNum - 5 * lng5
    
    strAns = strAns & String(lng1000, "M")
    strAns = strAns & String(lng500, "D")
    strAns = strAns & String(lng100, "C")
    strAns = strAns & String(lng50, "L")
    strAns = strAns & String(lng10, "X")
    strAns = strAns & String(lng5, "V")
    strAns = strAns & String(lngNum, "I")
    
    '例外
    strAns = Replace(strAns, "DCCCC", "CM")
    strAns = Replace(strAns, "LXXXX", "XC")
    strAns = Replace(strAns, "VIIII", "IX")
    strAns = Replace(strAns, "CCCC", "CD")
    strAns = Replace(strAns, "XXXX", "XL")
    strAns = Replace(strAns, "IIII", "IV")
    
    ToRoman = strAns

End Function
'--------------------------------------------------------------
'　文字列の分割（カンマ）
'--------------------------------------------------------------
Public Function CsvParser(ByVal strBuf As String, Optional ByVal c As String = vbTab) As Variant

    Dim lngLen As Long
    Dim lngCnt As Long
    Dim i As Long
    Dim strCol As String
    
    Dim blnSw As Boolean
    
    Const C_QUAT As String = """"
'    Const C_COMA As String = ","
    
    Dim Result() As Variant
    
    lngLen = Len(strBuf)
    blnSw = False
    strCol = ""
    lngCnt = 0
    
    For i = 1 To lngLen
    
        Dim strChar As String
        strChar = Mid$(strBuf, i, 1)
        
        Select Case strChar
            Case C_QUAT
                If blnSw Then
                    blnSw = False
                Else
                    blnSw = True
                End If
            Case c
                If blnSw Then
                    strCol = strCol & strChar
                Else
                    lngCnt = lngCnt + 1
                    ReDim Preserve Result(0 To lngCnt - 1)
                    Result(lngCnt - 1) = strCol
                    strCol = ""
                End If
            Case Else
                strCol = strCol & strChar
        End Select

    Next
    
    lngCnt = lngCnt + 1
    ReDim Preserve Result(0 To lngCnt - 1)
    Result(lngCnt - 1) = strCol

    CsvParser = Result
    
End Function
'--------------------------------------------------------------
'　文字列の分割（カンマ）
'--------------------------------------------------------------
Public Function SplitEx(ByVal strBuf As String, Optional ByVal c As String = vbTab) As Variant

    Dim lngLen As Long
    Dim lngCnt As Long
    Dim i As Long
    Dim strCol As String
    Dim blnSw As Boolean
    
    Const C_QUAT As String = """"
    
    Dim Result() As Variant
    
    lngLen = Len(strBuf)
    blnSw = False
    strCol = ""
    lngCnt = 0
    
    Dim IC As ICursor
    Set IC = Constructor(New CharCursor, strBuf)
    
    Do Until IC.Eof
    
        Dim blnBreak As Boolean
        blnBreak = False
        strCol = ""
        
        Do Until IC.Eof Or blnBreak
        
            If IC = C_QUAT And IC.PreviousItem <> "\" Then
                blnSw = Not blnSw
            End If
            
            If blnSw Then
                strCol = strCol & IC
            Else
                Select Case IC
                    Case c
                        blnBreak = True
                    Case Else
                        strCol = strCol & IC
                End Select
            End If
            IC.MoveNext
        Loop
    
        lngCnt = lngCnt + 1
        ReDim Preserve Result(0 To lngCnt - 1)
        Result(lngCnt - 1) = strCol
    
    Loop

    SplitEx = Result
    
End Function
Public Function StartsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    StartsWith = (strWith = Left$(strBuf, Len(strWith)))

End Function
Public Function EndsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    EndsWith = (strWith = Right$(strBuf, Len(strWith)))

End Function
