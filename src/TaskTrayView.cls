VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskTrayView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'>### TaskTrayView クラス
'>
'>**Remarks**
'>
'>タスクトレイアイコン表示＆メッセージ<br>
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
' TODO: 64bitでエラーにはならないがで表示されない？
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor

Private hWnd As LongPtr
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal obj As Object, hWnd As LongPtr) As LongPtr

Private Declare PtrSafe Function Shell_NotifyIconW Lib "shell32" (ByVal dwMessage As Long, pnid As NOTIFYICONDATAW) As Boolean
Private Declare PtrSafe Function ExtractIconExW Lib "shell32" (ByVal lpszFile As LongPtr, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As LongPtr)

Private Const CHAR_8  As Long = 8 - 1 '0 to 7
Private Const WCHAR_64  As Long = 64 * 2 - 1 '0 to 63
Private Const WCHAR_128 As Long = 128 * 2 - 1 '0 to 255
Private Const WCHAR_256 As Long = 256 * 2 - 1 '0 to 511

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(CHAR_8) As Byte
End Type

Private Type NOTIFYICONDATAW
    cbsize                  As Long
    hWnd                    As LongPtr
    uID                     As Long
    uFlags                  As Long
    uCallbackMessage        As Long
    hIcon                   As LongPtr
    szTip(WCHAR_128)        As Byte
    dwState                 As Long
    dwStateMask             As Long
    szInfo(WCHAR_256)       As Byte
    uTimeoutOrVersion       As Long
    szInfoTitle(WCHAR_64)   As Byte
    dwInfoFlags             As Long
    guidItem                As GUID
    hBalloonIcon            As LongPtr
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETVERSION = &H4
 
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIF_GUID = &H20

Private Const NIIF_LARGE_ICON = &H20
Private Const NIIF_INFO = &H1

Private Const NIIF_USER = &H4
Private Const NOTIFYICON_VERSION_4 = 4

Private typ_Icon As NOTIFYICONDATAW '自クラスで管理するアイコンのデータ(構造体)
'---------------------------------------------------------------
' コンストラクタ
'---------------------------------------------------------------
Private Function IConstructor_Instancing(Args As Collection) As Object

    Dim obj As Object
    Dim strMessage As String

    If Args.Count <> 2 Then
        Exit Function
    End If
    
    Set obj = Args(1)
    strMessage = Args(2)

    WindowFromAccessibleObject obj, hWnd
    AddIcon hWnd, strMessage
    
    Set IConstructor_Instancing = Me

End Function
'---------------------------------------------------------------
' デストラクタ
'---------------------------------------------------------------
Private Sub Class_Terminate()
    Dim typ_Dummy As NOTIFYICONDATAW
    Call DeleteIcon
    typ_Icon = typ_Dummy
End Sub
'---------------------------------------------------------------
' アイコン追加
'---------------------------------------------------------------
Private Sub AddIcon(ByVal hWnd As LongPtr, ByVal ToolTip As String)
 
    Dim IconLarge As LongPtr
    Dim IconSmall As LongPtr
    Dim IconIndex As Long
    Dim Ret As Boolean
 
    Dim strFile As String
    
    strFile = Application.Path & "\" & "EXCEL.EXE"
    IconIndex = 1
    
    Call ExtractIconExW(StrPtr(strFile), IconIndex, IconLarge, IconSmall, 1)

    With typ_Icon
        
        .cbsize = Len(typ_Icon)
        .hWnd = hWnd
        .uID = 0
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE Or NIF_INFO
        .hIcon = IconSmall

        ToolTip = Left$(ToolTip, 127) & vbNullChar
        CopyMemory .szTip(0), ByVal StrPtr(ToolTip), LenB(ToolTip)

    End With
     
    DrawMenuBar hWnd
     
    Ret = Shell_NotifyIconW(NIM_ADD, typ_Icon)
    
    If Ret Then
    Else
        MsgBox "Shell_NotifyIconW Faild", vbCritical
    End If
 
End Sub
'---------------------------------------------------------------
' バルーンメッセージ
'---------------------------------------------------------------
Public Sub ShowBalloon(ByVal Title As String, ByVal message As String, Optional ByVal BalloonTimeOutSecond As Long = 5)

    Dim IconLarge As LongPtr
    Dim IconSmall As LongPtr
    Dim IconIndex As Long
    Dim Ret As Boolean
 
    With typ_Icon
        
        .cbsize = Len(typ_Icon)
        .uFlags = NIF_INFO
        .uTimeoutOrVersion = BalloonTimeOutSecond * 1000 'タイムアウト時間(ms)
        .dwInfoFlags = NIIF_INFO
        
        message = Left$(message, 255) & vbNullChar
        CopyMemory .szInfo(0), ByVal StrPtr(message), LenB(message)
        
        Title = Left$(Title, 63) & vbNullChar
        CopyMemory .szInfoTitle(0), ByVal StrPtr(Title), LenB(Title)
        
    End With
    
    Ret = Shell_NotifyIconW(NIM_MODIFY, typ_Icon)
     
    If Ret Then
    Else
        MsgBox "Shell_NotifyIconW Faild", vbCritical
    End If

End Sub
'---------------------------------------------------------------
' アイコン削除
'---------------------------------------------------------------
Private Sub DeleteIcon()

    With typ_Icon
        .cbsize = Len(typ_Icon)
        .uFlags = NIF_ICON
        .uID = 0
    End With
     
    Call Shell_NotifyIconW(NIM_DELETE, typ_Icon) 'API実行
     
End Sub


