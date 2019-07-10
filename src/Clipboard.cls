VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [hidennotare] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/hidennotare
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
'   Lang.cls
'   Registry.cls
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成
' 2018-12-18 Ver.1.0.1 GetClipText 上限の 4096バイトを撤廃
' 2018-12-19 Ver.1.0.2 GetClipText で CF_UNICODETEXT に変更
'                      SetClipText  で CF_UNICODETEXT に変更
'                      GetFileNameFromCliをUnicode対応
'                      SetCopyClipTextをUnicode対応
'-----------------------------------------------------------------------------------------------------
Option Explicit

'クリップボード関数
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function DragQueryFileW Lib "shell32.dll" (ByVal hDrop As LongPtr, ByVal UINT As Long, ByVal lpszFile As LongPtr, ByVal ch As Long) As Long

'グローバルヒープ関数
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

'その他関数
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As LongPtr, ByRef IPic As IPicture) As Long
Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

Private Type PICTDESC
    cbSizeofStruct As Long
    picType        As Long
    hImage         As LongPtr
    Option1        As LongPtr
    Option2        As LongPtr
End Type
Private Type GUID
    Data1          As Long
    Data2          As Integer
    Data3          As Integer
    Data4(7)       As Byte
End Type

Private Const CF_TEXT As Long = 1
Private Const CF_HDROP As Long = 15
Private Const CF_BITMAP      As Long = 2
Private Const CF_PALETTE     As Long = 9
Private Const CF_UNICODETEXT As Long = 13

Private Const CC_RGBINIT = &H1                '色のデフォルト値を設定
Private Const CC_LFULLOPEN = &H2              '色の作成を行う部分を表示
Private Const CC_PREVENTFULLOPEN = &H4        '色の作成ボタンを無効にする
Private Const CC_SHOWHELP = &H8               'ヘルプボタンを表示

Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4
'**
' コピーアドレスの取得
'**
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
'--------------------------------------------------------------
'クリップボードをクリアする
'--------------------------------------------------------------
Public Sub ClearClipboard()

    If OpenClipboard(0&) <> 0 Then
        Call EmptyClipboard
        Call CloseClipboard
    End If

End Sub
'----------------------------------------------------------------------
' クリップボードのビットマップデータから Picture オブジェクトを作成
'----------------------------------------------------------------------
Public Function CreatePictureFromClipboard(o As Object) As StdPicture
  
    Dim hImg      As LongPtr
    Dim hPalette As LongPtr
    Dim hCopy As LongPtr
    
    Dim uPictDesc As PICTDESC
    Dim uGUID     As GUID
    
    Set CreatePictureFromClipboard = Nothing
  
    On Error GoTo e
    
    'クリップボードの保存
    If OpenClipboard(0&) <> 0 Then
        Call EmptyClipboard
        Call CloseClipboard
    End If
    
    '指定シェイプをビットマップでクリップボードに貼り付け
10  o.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

    Call CopyClipboardSleep
        
    If IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
    
        If OpenClipboard(0&) <> 0 Then
            
            hImg = GetClipboardData(CF_BITMAP)
        
            If hImg <> 0 Then
          
                hCopy = CopyImage(hImg, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                
                With uPictDesc
                    .cbSizeofStruct = Len(uPictDesc)
                    .picType = 1
                    .hImage = hCopy
                    .Option1 = 0&
                End With
                
                With uGUID
                    .Data1 = &H20400
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With
                
                Call OleCreatePictureIndirect(uPictDesc, uGUID, True, CreatePictureFromClipboard)
            
                Call EmptyClipboard
                
            End If
            
            Call CloseClipboard
        End If
        
    End If
    
    'クリップボードの復元
    Exit Function
e:
    If Erl = 10 Then
        Resume
    End If
    
End Function
'--------------------------------------------------------------
'　クリップボードからファイル名を取得
'--------------------------------------------------------------
Public Function GetCopyClipText() As String

    Dim hData As LongPtr
    Dim files As Long
    Dim i As Long
    Dim strFilePath As String
    Dim Ret As String
    Dim SB As StringBuilder
    
    Set SB = New StringBuilder
    
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
    
    GetCopyClipText = SB.ToJoin(vbCrLf)
    
End Function
'クリップボードにテキストデータを書き込むプロシージャ
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

  If blnErrflg Then Message.Error "クリップボードに情報が書き込めません"

End Sub
'クリップボードからテキストデータを取得するプロシージャ(Unicode対応)
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

'クリップボードにファイルコピー情報を書き込むプロシージャ
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
    
    If blnErrflg Then Message.Error "クリップボードに情報が書き込めません"

End Sub
'------------------------------------------------------------------------------------------------------------------------
' CopyPictureがJavaアプリやクリップボードツールなどで失敗する対策
'------------------------------------------------------------------------------------------------------------------------
Public Sub CopyClipboardSleep()
    DoEvents
    Sleep Val(Registry.GetSetting("Option", "ClipboardSleep", 0))
End Sub


