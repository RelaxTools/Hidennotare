VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICON"
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
'>### Icon クラス
'>
'>**Remarks**
'>
'>Iconの取得を行うクラスです。<br>
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-09-03|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"

Private Const PICTYPE_UNINITIALIZED As Long = (-1)
Private Const PICTYPE_NONE          As Long = 0
Private Const PICTYPE_BITMAP        As Long = 1
Private Const PICTYPE_METAFILE      As Long = 2
Private Const PICTYPE_ICON          As Long = 3
Private Const PICTYPE_ENHMETAFILE   As Long = 4
    
Private Type GUID
    Data1          As Long
    Data2          As Integer
    Data3          As Integer
    Data4(7)       As Byte
End Type
    
Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hpal As LongPtr
End Type

Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As Long, obj As Any) As Long
Private Declare PtrSafe Function ExtractIconExW Lib "Shell32" (ByVal lpszFile As LongPtr, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long

Private Const BOOL_TRUE      As Long = 1
Private Const BOOL_FALSE     As Long = 0

'Common HRESULT Values
Private Const S_OK              As Long = &H0        'Operation successful
Private Const E_ABORT           As Long = &H80004004 'Operation aborted
Private Const E_ACCESSDENIED    As Long = &H80070005 'General access denied error
Private Const E_FAIL            As Long = &H80004005 'Unspecified failure
Private Const E_HANDLE          As Long = &H80070006 'Handle that is not valid
Private Const E_INVALIDARG      As Long = &H80070057 'One or more arguments are not valid
Private Const E_NOINTERFACE     As Long = &H80004002 'No such interface supported
Private Const E_NOTIMPL         As Long = &H80004001 'Not implemented
Private Const E_OUTOFMEMORY     As Long = &H8007000E 'Failed to allocate necessary memory
Private Const E_POINTER         As Long = &H80004003 'Pointer that is not valid
Private Const E_UNEXPECTED      As Long = &H8000FFFF 'Unexpected failure

Private Const MAX_PATH                  As Long = 260 * 2 - 1
Private Const MAX_TYPE                  As Long = 80 * 2 - 1

Private Type SHFILEINFO
    hIcon As LongPtr
    iIcon As Long
    dwAttributes As Long
    szDisplayName(0 To MAX_PATH) As Byte
    szTypeName(0 To MAX_TYPE) As Byte
End Type

Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H0
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const SHGFI_ATTRIBUTES = &H800
Const SHGFI_DISPLAYNAME = &H200
Const SHGFI_EXETYPE = &H2000
Const SHGFI_ICON = &H100
Const SHGFI_ICONLOCATION = &H1000
Const SHGFI_LARGEICON = &H0
Const SHGFI_LINKOVERLAY = &H8000
Const SHGFI_OPENICON = &H2
Const SHGFI_PIDL = &H8
Const SHGFI_SELECTED = &H10000
Const SHGFI_SHELLICONSIZE = &H4
Const SHGFI_SMALLICON = &H1
Const SHGFI_SYSICONINDEX = &H4000
Const SHGFI_TYPENAME = &H400
Const SHGFI_USEFILEATTRIBUTES = &H10

Private Declare PtrSafe Function SHGetFileInfoW Lib "Shell32" (ByVal pszPath As LongPtr, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As LongPtr

Private Declare PtrSafe Function ExtractAssociatedIconW Lib "Shell32" (ByVal hInst As LongPtr, ByVal lpIconPath As LongPtr, lpiIcon As Long) As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "USER32" (ByVal hIcon As LongPtr) As Long

Public Enum IconConstants
    IconSmall = 1
    IconLarge
End Enum

Private hIcon As LongPtr
Private mFileName As String
Private mSize As IconConstants
'--------------------------------------------------------------
' Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    
    hIcon = 0
    
End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = Icon.CreateObject(FileName,
'>                            [size = IconConstants.IconSmall])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を設定|
'>|size|省略可|IconConstants|アイコンサイズの設定(IconSmall/IconLarge)。デフォルト小(16×16)|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Icon|Iconオブジェクト|
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>指定したファイル名のアイコンを取得してピクチャ形式に変換します。<br>
'>
'>**Sample**
'>
'>```
'>
'>    'ImageList に ファイルのアイコンを設定する例
'>
'>    Dim Key As String
'>    Dim Ext As String
'>
'>    Ext =  UCase(FileIO.GetExtensionName(FileName))
'>
'>    If Ext = "EXE" Then
'>        Key = FileIO.GetFileName(FileName)
'>    Else
'>        Key = Ext
'>    End If
'>
'>    ImageList1.ListImages.Add , Key, Icon.CreateObject(FileName).ToPicture
'>
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function CreateObject(ByVal FileName As String, Optional ByVal Size As IconConstants) As ICON
    
    Dim IC As ICON
    
    Set IC = New ICON
    
    IC.FileName = FileName
    IC.Size = Size
    
    IC.Init
    
    Set CreateObject = IC
    
End Function
Public Property Let FileName(v As String)
    mFileName = v
End Property
Public Property Let Size(v As IconConstants)
    mSize = v
End Property
Public Sub Init()
    
    Dim hIconLarge As LongPtr
    Dim hIconSmall As LongPtr
    Dim FileName As String
    Dim Size As IconConstants
    
    FileName = mFileName
    Size = mSize
    
    If UCase(Right$(FileName, 4)) = ".EXE" Then
        
        ExtractIconExW StrPtr(FileName), 1, hIconLarge, hIconSmall, 1

        Select Case Size
            Case IconConstants.IconSmall
                hIcon = hIconSmall
            
            Case IconConstants.IconLarge
                hIcon = hIconLarge
            
            Case Else
                hIcon = 0
        
        End Select
    
    End If

    Dim info  As SHFILEINFO
    Dim ret As LongPtr

    If hIcon = 0 Then
        
        If Size = IconConstants.IconSmall Then
            ret = SHGetFileInfoW(StrPtr(FileName), FILE_ATTRIBUTE_NORMAL, info, Len(info), SHGFI_USEFILEATTRIBUTES Or SHGFI_TYPENAME Or SHGFI_ICON Or SHGFI_SMALLICON)
        Else
            ret = SHGetFileInfoW(StrPtr(FileName), FILE_ATTRIBUTE_NORMAL, info, Len(info), SHGFI_USEFILEATTRIBUTES Or SHGFI_TYPENAME Or SHGFI_ICON)
        End If
        
        hIcon = info.hIcon
    End If

End Sub
'--------------------------------------------------------------
' Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    If hIcon <> 0 Then
        Call DestroyIcon(hIcon)
    End If
End Sub
'>---
'>#### ToPicture(Icon)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToString()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|Icon|Iconオブジェクトを指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|IPicture|VBAで使用可能なピクチャ形式|
'>
'>**Remarks**
'>
'>ピクチャ変換メソッドです。<br>
'>VBAで使用可能なピクチャ形式に変換します。<br>
'>
'>**See also**
'>
'>* None
'>
Function ToPicture() As IPictureDisp
    
    Dim uPicInfo As PICTDESC
    Dim result As Long
    
    If hIcon = 0 Then
        Err.Raise result, TypeName(Me) & ".ToPicture", "hIcon = null "
    End If
    
    With uPicInfo
      .Size = LenB(uPicInfo)
      .Type = PICTYPE_ICON
      .hPic = hIcon
      .hpal = 0&
    End With
        
    result = OleCreatePictureIndirect(uPicInfo, ConvGUID(IID_IPictureDisp), BOOL_TRUE, ToPicture)
    If result <> S_OK Then
        Err.Raise result, TypeName(Me) & ".ToPicture", "OleCreatePictureIndirect error"
    End If

End Function
Private Function ConvGUID(ByVal strGUID As String) As GUID
    
    Call IIDFromString(StrPtr(strGUID), ConvGUID)

End Function


