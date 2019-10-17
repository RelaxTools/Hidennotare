VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Core"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'>### Core クラス
'>
'>**Remarks**
'>
'>Coreクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>基本的な機能を保有する。<br>
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Enum NewLineCodeConstants
    NewLineCodeCRLF
    NewLineCodeLF
    NewLineCodeNone
End Enum

Public Enum EncodeConstants
    EncodeCP932
    EncodeUTF8
    EncodeUTF16LE
End Enum

Public Enum OpenModeConstants
    OpenModeOutput
    OpenModeAppend
End Enum

Public Enum CsvReaderHeaderConstants
    UseFirstRow
    AddHedder
    HeaderNone
End Enum
'-----------------------------------------------------------------------------------------------------
' 配列またはParamArrayをCollectionに変換
'-----------------------------------------------------------------------------------------------------
Public Function ArrayToCollection(Args As Variant) As Collection
        
    Dim col As Collection
    Dim v As Variant
    
    Set col = New Collection
        
    If UBound(Args) = 0 And IsArray(Args) Then
        For Each v In Args(0)
            col.Add v
        Next
    Else
        For Each v In Args
            col.Add v
        Next
    End If

    Set ArrayToCollection = col

End Function
'-----------------------------------------------------------------------------------------------------
' VBA 個人的汎用処理 https://qiita.com/nukie_53/items/bde16afd9a6ca789949d
' @nukie_53
' Set/Letを隠蔽するプロパティ
'-----------------------------------------------------------------------------------------------------
Public Property Let SetVar(outVariable As Variant, inExpression As Variant)
    
    Select Case True
        Case VBA.IsObject(inExpression), VBA.VarType(inExpression) = vbDataObject
            
            Set outVariable = inExpression
        
        Case Else
            
            Let outVariable = inExpression
    
    End Select

End Property
'---------------------------------------------------------------------------------------------------
' Dictionary判定
'---------------------------------------------------------------------------------------------------
Public Function IsDictionary(v As Variant) As Boolean

    IsDictionary = True
    
    Select Case TypeName(v)
        Case "Dictionary"
        Case "OrderedDictionary"
        Case "SortedDictionary"
        Case Else
            IsDictionary = False
    End Select

End Function
'---------------------------------------------------------------------------------------------------
' List判定
'---------------------------------------------------------------------------------------------------
Public Function IsList(v As Variant) As Boolean

    IsList = True
    
    Select Case TypeName(v)
        Case "ArrayList"
        Case "LinkedList"
        Case "Collection"
        Case Else
            IsList = False
    End Select

End Function
'>---
'>#### TryUbound
'>
'>**Syntax**
'>
'>```
'>flg = Core.TryUBound(ary, idx, [dimension = 1])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|ary|必須|Variant|チェックをする配列を指定|
'>|idx|必須|Long|配列の上限を返却する変数を指定|
'>|dimension|省略可|Long|省略時１（次元）。テストする次元数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|正常:True / エラー:False|
'>
'>**Remarks**
'>
'>エラー検知付きUBound<br>
'>テストする配列が１次元の場合に２次元の上限を求めようとすると実行時エラーになる。<br>
'>そのためエラー検知可能なUboundとして作成。エラーの場合には False が返却される。
'>
'>**Example**
'>
'>**See also**
'>
'>* Core.TryLBound
'>
Public Function TryUBound(ary As Variant, idx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    idx = UBound(ary, dimension)

    TryUBound = True
    
    Exit Function
e:
    TryUBound = False

End Function
'>---
'>#### TryLBound
'>
'>**Syntax**
'>
'>```
'>flg = Core.TryLBound(ary, idx, [dimension = 1])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|ary|必須|Variant|チェックをする配列を指定|
'>|idx|必須|Long|配列の下限を返却する変数を指定|
'>|dimension|省略可|Long|省略時１（次元）。テストする次元数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|正常:True / エラー:False|
'>
'>**Remarks**
'>
'>エラー検知付きUBound<br>
'>テストする配列が１次元の場合に２次元の下限を求めようとすると実行時エラーになる。<br>
'>そのためエラー検知可能なUboundとして作成。エラーの場合には False が返却される。
'>
'>**Example**
'>
'>**See also**
'>
'>* Core.TryUBound
'>
Public Function TryLBound(ary As Variant, lngIdx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    lngIdx = LBound(ary, dimension)

    TryLBound = True
    
    Exit Function
e:
    TryLBound = False

End Function
'------------------------------------------------------------------------------------------------------------------------
' 上位バイト取得
'------------------------------------------------------------------------------------------------------------------------
Public Function UByte(ByVal lngValue As Long) As Long
    UByte = RShift((lngValue And &HFF00&), 8)
End Function
'------------------------------------------------------------------------------------------------------------------------
' 下位バイト取得
'------------------------------------------------------------------------------------------------------------------------
Public Function LByte(ByVal lngValue As Long) As Long
    LByte = lngValue And &HFF&
End Function
'------------------------------------------------------------------------------------------------------------------------
' 左シフト
'------------------------------------------------------------------------------------------------------------------------
Public Function LShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    LShift = lngValue * (2 ^ lngKeta)
End Function
'------------------------------------------------------------------------------------------------------------------------
' 右シフト
'------------------------------------------------------------------------------------------------------------------------
Public Function RShift(ByVal lngValue As Long, ByVal lngKeta As Long) As Long
    RShift = lngValue \ (2 ^ lngKeta)
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
'>---
'>#### Escape(Core)
'>
'>**Syntax**
'>
'>```
'>ret = Core.Escape(strBuf)
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
    Dim SB As StringBuilder
    Dim i As Long
    
    Set SB = New StringBuilder
    
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
'>#### Unescape(Core)
'>
'>**Syntax**
'>
'>```
'>ret = Core.Unescape(strBuf)
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
    Dim SB As StringBuilder
    Dim i As Long
    
    Set SB = New StringBuilder
    i = 1
    Do Until i > Len(v)

        strChr = Mid$(v, i, 2)
        
        Select Case strChr
            Case "\u"
            
                If i + 6 - 1 > Len(v) Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Unescape", "format error"
                End If
            
                SB.Append ChrW$(Val("&H" & Mid$(v, i + 2, 4)))
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
