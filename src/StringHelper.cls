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
'>### StringHelper クラス
'>
'>**Remarks**
'>
'>文字列編集<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* CharCursor
'>* StringBuilder
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
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
    
        lngBuf = Core.LShift(bytBuf(i + 1), 8) + bytBuf(i)
    
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
''----------------------------------------------------------------------------------
''　文字列の右端から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
''----------------------------------------------------------------------------------
'Public Function AscRight(ByVal var As Variant, ByVal lngSize As Long) As String
'
'    Dim ascVar As Variant
'
'    ascVar = VBA.StrConv(var, vbFromUnicode)
'
'    AscRight = VBA.StrConv(RightB(ascVar, lngSize), vbUnicode)
'
'End Function
''----------------------------------------------------------------------------------
''　文字列から指定した文字数分の文字列を返す。漢字２バイト、半角１バイト。
''----------------------------------------------------------------------------------
'Public Function AscMid(ByVal var As Variant, ByVal lngPos As Long, Optional ByVal varSize As Variant) As String
'
'    Dim ascVar As Variant
'
'    ascVar = VBA.StrConv(var, vbFromUnicode)
'
'    If IsMissing(varSize) Then
'        AscMid = VBA.StrConv(MidB(ascVar, lngPos), vbUnicode)
'    Else
'        AscMid = VBA.StrConv(MidB(ascVar, lngPos, varSize), vbUnicode)
'    End If
'
'End Function
''--------------------------------------------------------------
''　文字列の分割（カンマ）
''--------------------------------------------------------------
'Public Function SplitEx(ByVal strBuf As String, Optional ByVal Separator As String = vbTab) As Variant
'
'    Dim lngLen As Long
'    Dim lngCnt As Long
'    Dim i As Long
'    Dim strCol As String
'    Dim blnSw As Boolean
'
'    Const C_QUAT As String = """"
'
'    Dim result() As Variant
'
'    lngLen = Len(strBuf)
'    blnSw = False
'    strCol = ""
'    lngCnt = 0
'
'    Dim IC As ICursor
'    Set IC = Constructor(New CharCursor, strBuf)
'
'    Do Until IC.Eof
'
'        Dim blnBreak As Boolean
'        blnBreak = False
'        strCol = ""
'
'        Do Until IC.Eof Or blnBreak
'
'            If IC.Item = C_QUAT And IC.Item(-1) <> "\" Then
'                blnSw = Not blnSw
'            End If
'
'            If blnSw Then
'                strCol = strCol & IC
'            Else
'                Select Case IC
'                    Case Separator
'                        blnBreak = True
'                    Case Else
'                        strCol = strCol & IC
'                End Select
'            End If
'            IC.MoveNext
'        Loop
'
'        lngCnt = lngCnt + 1
'        ReDim Preserve result(0 To lngCnt - 1)
'        result(lngCnt - 1) = strCol
'
'    Loop
'
'    SplitEx = result
'
'End Function
'--------------------------------------------------------------
' CSVパーサー RFC4180準拠
' コーテーションの有無, デリミタ指定あり
'--------------------------------------------------------------
Public Function CsvParser(ByVal strBuf As String, Optional ByVal Quatation As Boolean = False, Optional ByVal Delimiter As String = ",") As Collection

    Const C_QUAT As String = """"
    Dim IC As ICursor
    Dim sw As Boolean
    Dim blnLineBreak As Boolean
    Dim blnItemBreak As Boolean
    Dim col As Collection
    Dim Row As Collection
    Dim lngQuot As Long
    
    lngQuot = 0
    
    sw = False
    
    Set Row = New Collection
    
    Set IC = Constructor(New CharCursor, strBuf)
    Do Until IC.Eof
        
        '初期化
        Set col = New Collection
        blnLineBreak = False

        Do Until IC.Eof Or blnLineBreak
        
            '初期化
            Dim SB As StringBuilder
            Set SB = New StringBuilder
            blnItemBreak = False
            
            Do Until IC.Eof Or blnLineBreak Or blnItemBreak

                Select Case IC.Item
                    Case C_QUAT

                        'コーテーションありの場合
                        If Quatation Then

                            lngQuot = lngQuot + 1

                            If sw Then

                                '次の文字がQuatation
                                If IC.Item(1) = C_QUAT Then
                                    lngQuot = lngQuot + 1
                                    SB.Append C_QUAT
                                    IC.MoveNext
                                Else
                                    sw = False
                                End If

                            Else
                                sw = True
                            End If
                        End If

                    Case vbCr

                        If sw Then
                            SB.Append IC.Item
                        End If

                    Case vbLf

                        If sw Then
                            SB.Append IC.Item
                        Else
                            blnLineBreak = True
                        End If

                    Case Delimiter

                        If sw Then
                            SB.Append IC.Item
                        Else
                            blnItemBreak = True
                        End If

                    Case Else

                        SB.Append IC.Item

                End Select
            
                IC.MoveNext
            Loop
            
            '列追加処理
            col.Add SB.ToString
            
        Loop
        
        '行追加処理
        Row.Add col
        
    Loop
    
    'ダブルコーテーションが偶数ではない場合
    If lngQuot Mod 2 <> 0 Then
        Err.Raise 512 + 1, TypeName(Me) & ".CsvParser", "Invalid Format(Quotation)"
    End If
    
    Set CsvParser = Row
    
End Function
'--------------------------------------------------------------
'　開始文字列の判定
'--------------------------------------------------------------
Public Function StartsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    StartsWith = (strWith = Left$(strBuf, Len(strWith)))

End Function
'--------------------------------------------------------------
'　終了文字列の判定
'--------------------------------------------------------------
Public Function EndsWith(ByVal strBuf As String, ByVal strWith As String) As Boolean

    EndsWith = (strWith = Right$(strBuf, Len(strWith)))

End Function
