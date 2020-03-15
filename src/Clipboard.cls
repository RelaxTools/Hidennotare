VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clipboard"
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
'>### Clipboard クラス
'>
'>**Remarks**
'>
'>Clipboard に関するクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* Registry
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>|2018-12-18|Ver.1.0.1|GetClipText 上限の 4096バイトを撤廃|
'>|2018-12-19|Ver.1.0.2|GetClipText CF_UNICODETEXT に変更<br>SetClipText CF_UNICODETEXT に変更<br>GetFileNameFromCliをUnicode対応<br>SetCopyClipTextをUnicode対応|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

'クリップボード関数
Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "USER32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function DragQueryFileW Lib "shell32.dll" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpszFile As LongPtr, ByVal ch As Long) As Long

'グローバルヒープ関数
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

'その他関数
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

Private Const CF_TEXT As Long = 1
Private Const CF_HDROP As Long = 15
Private Const CF_BITMAP      As Long = 2
Private Const CF_PALETTE     As Long = 9
Private Const CF_UNICODETEXT As Long = 13
'>---
'>#### getObjectLink(Clipboard)
'>
'>**Syntax**
'>
'>```
'>ret = Clipboard.getObjectLink()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|Excelのコピー元アドレス|
'>
'>**Remarks**
'>
'>コピー元アドレスの取得<br>
'>Excelでコピーした時のハイライトの範囲を取得できる。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Function getObjectLink() As String

    Dim p As LongPtr
    Dim hMem As LongPtr
    Dim lngDataLen As LongPtr
    Dim lngRet As LongPtr

    Const MAXSIZE = 4096
    Dim MyString As String
    Dim size As Long
    Dim data() As Byte
    Dim i As Long
  
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
    
        hMem = GetClipboardData(RegisterClipboardFormat("ObjectLink"))
        
        If hMem <> 0 Then
        
            size = CLng(GlobalSize(hMem))
            p = GlobalLock(hMem)
            
            If p <> 0 Then
                
                ReDim data(0 To size - 1)
                Call CopyMemory(data(0), ByVal p, size)
                
                Call GlobalUnlock(hMem)
                
                For i = 0 To size - 1
                    If data(i) = 0 Then
                        data(i) = &H9
                    End If
                Next i
                MyString = StrConv(data(), vbUnicode)
                
            End If
        
        End If
        
        CloseClipboard
    
    End If
    
    getObjectLink = MyString

End Function
'>---
'>#### ClearClipboard(Clipboard)
'>
'>**Syntax**
'>
'>```
'>String = Clipboard.ClearClipboard()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>クリップボードのクリア<br>
'>クリップボードをクリアします。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub ClearClipboard()

    If OpenClipboard(0&) <> 0 Then
        Call EmptyClipboard
        Call CloseClipboard
    End If

End Sub

'>---
'>#### GetCopyClipText(Clipboard)
'>
'>**Syntax**
'>
'>```
'>ret = Clipboard.GetCopyClipText()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|クリップボードの改行(vbCrLf)で区切られたファイル情報を返却|
'>
'>**Remarks**
'>
'>エクスプローラコピー情報取得<br>
'>エクスプローラからのコピー情報を<br>
'>
'>**Example**
'>
'>```
'>Dim strBuf As String
'>Dim v As Variant
'>
'>strBuf = Clipboard.GetCopyClipText()
'>
'>v = Split(strBuf, vbCrLf)
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function GetCopyClipText() As String

    Dim hData As LongPtr
    Dim files As Long
    Dim i As Long
    Dim strFilePath As String
    Dim ret As String
    Dim SB As IStringBuilder
    
    Set SB = StringBuilder.NewInstance
    
    If OpenClipboard(0) <> 0 Then
   
        hData = GetClipboardData(CF_HDROP)
        
        If Not IsNull(hData) Then
            
            'ファイルの数を取得
            files = DragQueryFileW(hData, -1, 0, 0)
            For i = 0 To files - 1 Step 1
                
                'サイズを取得
                Dim lngSize As Long
                lngSize = DragQueryFileW(hData, i, 0, 0)
                
                'DragQueryFileWの返却するサイズは終端を含まない
                strFilePath = String$(lngSize + 1, vbNullChar)
                
                lngSize = DragQueryFileW(hData, i, StrPtr(strFilePath), Len(strFilePath))
                
                SB.Append Left$(strFilePath, lngSize)
            
            Next
        End If
        Call CloseClipboard
    
    End If
    
    GetCopyClipText = SB.ToString(vbCrLf)
    
End Function
'>---
'>#### SetClipText(Clipboard)
'>
'>**Syntax**
'>
'>```
'>Clipboard.SetClipText strData
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strData|必須|String|改行で区切られたファイル情報|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テキスト書き込み<br>
'>クリップボードにテキストデータを書き込むプロシージャ<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub SetClipText(ByVal strData As String)

  Dim p As LongPtr
  Dim hMem As LongPtr
  Dim lngDataLen As LongPtr
  Dim blnErrflg As Boolean
  Const GMEM_MOVEABLE = 2

  blnErrflg = True
  
  'クリップボードをオープン
  If OpenClipboard(0&) <> 0 Then
  
    'クリップボードを空にする
    If EmptyClipboard() <> 0 Then
    
        '終端文字を付加
        strData = strData & vbNullChar
    
        'グローバルメモリに書き込む領域を確保してそのハンドルを取得
        lngDataLen = LenB(strData)
        
        hMem = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
        
        If hMem <> 0 Then
      
            'グローバルメモリをロックしてそのポインタを取得
            p = GlobalLock(hMem)
            
            If p <> 0 Then
        
                '書き込むテキストをグローバルメモリにコピー
                Call CopyMemory(ByVal p, ByVal StrPtr(strData), lngDataLen)
                
                'グローバルメモリブロックのロックを解除
                Call GlobalUnlock(hMem)
                
                If SetClipboardData(CF_UNICODETEXT, hMem) <> 0 Then
                    blnErrflg = False
                End If

            End If
        End If
    End If
    'クリップボードをクローズ(これはWindowsに制御が
    '戻らないうちにできる限り速やかに行う)
    Call CloseClipboard
  End If

  If blnErrflg Then Err.Raise 520, TypeName(Me) & ".SetClipText", "not write clipboard."

End Sub
'>---
'>#### GetClipText(Clipboard)
'>
'>**Syntax**
'>
'>```
'>ret = Clipboard.GetClipText()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|クリップボードのテキストデータを返却|
'>
'>
'>**Remarks**
'>
'>テキストデータ取得<br>
'>クリップボードからテキストデータを取得するプロシージャ(Unicode対応)<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function GetClipText() As String

    Dim p As LongPtr
    Dim hMem As LongPtr
    Dim lngDataLen As LongPtr
    Dim lngRet As LongPtr
    Dim bytBuf() As Byte

    Dim MyString As String
    Dim blnUnicode As Boolean
  
    blnUnicode = False
    
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
    
        hMem = GetClipboardData(CF_UNICODETEXT)
        If hMem <> 0 Then
        
            p = GlobalLock(hMem)
            
            If p <> 0 Then
            
                lngDataLen = GlobalSize(hMem)
                ReDim bytBuf(0 To CLng(lngDataLen - 1))
                
                Call CopyMemory(bytBuf(0), ByVal p, lngDataLen)
                
                Call GlobalUnlock(hMem)
                
                MyString = bytBuf
                MyString = Mid$(MyString, 1, InStr(MyString, vbNullChar) - 1)
            
            End If
        
        End If
        
        lngRet = CloseClipboard()
    
    End If
    
    GetClipText = MyString

End Function
'>---
'>#### SetCopyClipText(Clipboard)
'>
'>**Syntax**
'>
'>```
'>Clipboard.SetCopyClipText strBuf()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strBuf()|必須|String配列|ファイル情報の配列|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル情報書き込み<br>
'>クリップボードにファイル情報を書き込むプロシージャ<br>
'>エクプローラのコピーやアウトルックの添付が可能<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub SetCopyClipText(strBuf() As String)

    Dim p As LongPtr
    Dim hMem As LongPtr
    Dim lngDataLen As LongPtr

    Dim blnErrflg As Boolean
    Const GMEM_MOVEABLE = 2
    
    Dim df As DROPFILES
    
    Dim strData As String
    Dim i As Long
    strData = ""
    For i = LBound(strBuf) To UBound(strBuf)
        strData = strData & strBuf(i) & vbNullChar
    Next
    strData = strData & vbNullChar

    blnErrflg = True
  
    'クリップボードをオープン
    If OpenClipboard(0&) <> 0 Then
  
        'クリップボードを空にする
        If EmptyClipboard() <> 0 Then
    
            'グローバルメモリに書き込む領域を確保してそのハンドルを取得
            lngDataLen = LenB(strData) + LenB(df)
            
            hMem = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
            
            If hMem <> 0 Then
            
                'グローバルメモリをロックしてそのポインタを取得
                p = GlobalLock(hMem)
                
                If p <> 0 Then
                
                    df.pFiles = LenB(df)
                    df.fWide = True 'UNICODE
            
                    '書き込むテキストをグローバルメモリにコピー
                    CopyMemory ByVal p, df, LenB(df)
                    CopyMemory ByVal (p + LenB(df)), ByVal StrPtr(strData), LenB(strData)
                    
                    'クリップボードにメモリブロックのデータを書き込み
                    If SetClipboardData(CF_HDROP, hMem) <> 0 Then
                        blnErrflg = False
                    End If
                
                    'グローバルメモリブロックのロックを解除
                    Call GlobalUnlock(hMem)
                    
                End If
                
            End If
            
            
            'テキストも一緒に書き込んでおく
            strData = ""
            For i = LBound(strBuf) To UBound(strBuf)
                strData = strData & strBuf(i) & vbCrLf
            Next
            strData = strData & vbNullChar
            
            'グローバルメモリに書き込む領域を確保してそのハンドルを取得
            lngDataLen = LenB(strData)
            
            hMem = GlobalAlloc(GMEM_MOVEABLE, lngDataLen)
            
            If hMem <> 0 Then
            
                'グローバルメモリをロックしてそのポインタを取得
                p = GlobalLock(hMem)
                
                If p <> 0 Then
            
                    '書き込むテキストをグローバルメモリにコピー
                    Call CopyMemory(ByVal p, ByVal StrPtr(strData), lngDataLen)
                    
                    'グローバルメモリブロックのロックを解除
                    Call GlobalUnlock(hMem)
    
                    If SetClipboardData(CF_UNICODETEXT, hMem) <> 0 Then
                        blnErrflg = False
                    End If

                End If
            End If
            
        End If
        
        'クリップボードをクローズ(これはWindowsに制御が
        '戻らないうちにできる限り速やかに行う)
        Call CloseClipboard
    End If
    
    If blnErrflg Then Err.Raise 520, TypeName(Me) & ".SetCopyClipText", "not write clipboard."

End Sub



