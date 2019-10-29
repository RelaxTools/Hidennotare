VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GDIPlus"
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
'>### GDIPlus クラス
'>
'>**Remarks**
'>
'>グラフィクス(GDI+)の処理をまとめたクラスです。<br>
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
'>|2019-08-26|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IGDIPlusParam

'クラスID
Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Const CLSID_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_JPEG As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_QUALITY As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Const PICTYPE_BITMAP As Long = 1
    
Private Type GUID
    Data1          As Long
    Data2          As Integer
    Data3          As Integer
    Data4(7)       As Byte
End Type
    
Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hBitmap As LongPtr, ByVal hPalet As LongPtr, ByRef bitmap As LongPtr) As Long

Private Declare PtrSafe Function GdipCreateBitmapFromHICON Lib "GDIPlus" (ByVal hIcon As LongPtr, ByRef bitmap As LongPtr) As Long

Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As LongPtr) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As PICTDESC, ByRef refiid As GUID, ByVal fPictureOwnsHandle As Long, obj As Any) As Long

Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hpal As LongPtr
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    GUID           As GUID
    NumberOfValues As Long
    Type           As Long
    Value          As LongPtr
End Type

Private Type EncoderParameters
    Count         As Long
    Parameter(15) As EncoderParameter
End Type

Private Declare PtrSafe Function GdipCreateSolidFill Lib "GDIPlus" (ByVal pColor As Long, ByRef brush As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "GDIPlus" (ByVal Image As LongPtr, Graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "GDIPlus" (ByVal Graphics As LongPtr, ByVal brush As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipSetSmoothingMode Lib "GDIPlus" (ByVal mGraphics As LongPtr, ByVal mSmoothingMode As Long) As Long
Private Declare PtrSafe Function GdipDeleteBrush Lib "GDIPlus" (ByVal mBrush As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "GDIPlus" (ByVal Graphics As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function GdipCreatePen1 Lib "GDIPlus" (ByVal pColor As Long, ByVal Width As Long, ByVal unit As Long, ByRef hPen As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawRectangle Lib "GDIPlus" (ByVal hGraphics As LongPtr, ByVal hPen As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipDeletePen Lib "GDIPlus" (ByVal hPen As LongPtr) As Long
Private Declare PtrSafe Function ExtractIconExW Lib "Shell32" (ByVal lpszFile As LongPtr, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long

Private Declare PtrSafe Function GdipFillEllipse Lib "GDIPlus" (ByVal hGraphics As LongPtr, ByVal brush As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipDrawEllipse Lib "GDIPlus" (ByVal hGraphics As LongPtr, ByVal hPen As LongPtr, ByVal X As Single, ByVal Y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long

Private Declare PtrSafe Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As LongPtr, ByVal FileName As LongPtr, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
    
Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Enum DrawStringConstants
    DrawStringFontStyleRegular = 0
    DrawStringFontStyleBold = 1
    DrawStringFontStyleItalic = 2
    DrawStringFontStyleUnderLine = 4
    DrawStringFontStyleStrikeThrough = 8
End Enum
    
Private Declare PtrSafe Function GdipCreateFromHDC Lib "GDIPlus" (ByVal hDc As LongPtr, ByRef Graphics As LongPtr) As Long
    
Private Declare PtrSafe Function GdipCreateFontFamilyFromName Lib "GDIPlus" (ByVal FontName As LongPtr, ByVal fontCollection As Long, fontFamily As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateFont Lib "GDIPlus" (ByVal fontFamily As LongPtr, ByVal emSize As Single, ByVal style As Integer, ByVal unit As Integer, font As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawString Lib "GDIPlus" (ByVal Graphics As LongPtr, ByVal str As LongPtr, ByVal Length As Long, ByVal font As LongPtr, layoutRect As RECTF, ByVal StringFormat As LongPtr, ByVal brush As LongPtr) As Long
    
Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter
    StringAlignmentFar
End Enum

Private Const LANG_NEUTRAL As Integer = 0
    
Private Declare PtrSafe Function GdipSetStringFormatAlign Lib "GDIPlus" (ByVal hFormat As LongPtr, ByVal align As Long) As Long
Private Declare PtrSafe Function GdipSetStringFormatLineAlign Lib "GDIPlus" (ByVal hFormat As LongPtr, ByVal align As Long) As Long

Private Declare PtrSafe Function GdipCreateStringFormat Lib "GDIPlus" (ByVal formatAttributes As Long, ByVal LANGID As Integer, ByRef hFormat As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteStringFormat Lib "GDIPlus" (ByVal hFormat As LongPtr) As Long

Private hGdiPlus As LongPtr
Private hGdiImage As LongPtr

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

Private Const SmoothingModeAntiAlias    As Long = &H4

Public Enum GDIPlusConstants
    GDIPlusCreateFromFile = 1
    GDIPlusCreateFromIconHandle
    GDIPlusCreateFromShape
    GDIPlusCreateFromIPictureDisp
    GDIPlusCreateImage
End Enum

Private Enum Status
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum

'Picrureオブジェクト操作
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_COPYRETURNORG As Long = &H4
Private Declare PtrSafe Function CopyImage Lib "USER32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr

'クリップボード関数
Private Declare PtrSafe Function OpenClipboard Lib "USER32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "USER32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "USER32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
Private Const CF_BITMAP      As Long = 2
Private Const CF_PALETTE     As Long = 9

Private Declare PtrSafe Function GetDC Lib "USER32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "USER32" (ByVal hWnd As LongPtr, ByVal hDc As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Const MAX_PATH                  As Long = 260 * 2 - 1
Private Const MAX_TYPE                  As Long = 80 * 2 - 1

Private Type SHFILEINFO
    hIcon As LongPtr
    iIcon As LongPtr
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
'Private Declare PtrSafe Function ExtractIconExW Lib "shell32" (ByVal lpszFile As LongPtr, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long

Private mEditType As GDIPlusConstants
Private mFileName As String
Private mhIcon As LongPtr
Private mWidth As Long
Private mHeight As Long
Private mColor As Long
Private mPic As IPictureDisp
Private mShape As Object
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = GDIPlus.CreateObject(GDIPlusConstants.GDIPlusCreateFile,
'>                               FileName)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|mode|必須|GDIPlusConstants|GDIPlusCreateFileを指定|
'>|FileName|必須|String|ファイル名を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(GDIPlus)|
'>
'>**Remarks**
'>
'>画像ファイルを読み込みます。(BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF)
'>
'>**Example**
'>
'>
'>```
'>Dim strFile As String
'>Dim GP As GDIPlus
'>
'>strFile = "C:\image.png"
'>
'>Set GP = GDIPlus.CreateObject(
'>                     GDIPlusConstants.GDIPlusCreateFile,
'>                     strFile)
'>
'>Image1.Picture = GP.ToPicture
'>
'>
'>```
'>
'>**See also**
'>
'>* None
'>
'--------------------------------------------------------------
Public Function CreateObjectFromFile(ByVal FileName As String) As IGDIPlus
                
    Dim GP As IGDIPlusParam
                
    Set GP = New GDIPlus
    
    GP.Initialize
    
    GP.EditType = GDIPlusConstants.GDIPlusCreateFromFile
    GP.FileName = FileName
    
    GP.Run
                
    Set CreateObjectFromFile = GP

End Function
Public Function CreateObjectFromIconHandle(ByVal hIcon As LongPtr) As IGDIPlus
                
    Dim GP As IGDIPlusParam
                
    Set GP = New GDIPlus
    
    GP.Initialize
    
    GP.EditType = GDIPlusConstants.GDIPlusCreateFromIconHandle
    GP.hIcon = hIcon
    
    GP.Run
                
    Set CreateObjectFromIconHandle = GP

End Function
Public Function CreateObjectFromShape(obj As Object) As IGDIPlus
                
    Dim GP As IGDIPlusParam
                
    Set GP = New GDIPlus
    
    GP.Initialize
    
    GP.EditType = GDIPlusConstants.GDIPlusCreateFromShape
    Set GP.Shape = obj
    
    GP.Run
                
    Set CreateObjectFromShape = GP
End Function
Public Function CreateObjectFromIPictureDisp(Pic As IPictureDisp) As IGDIPlus

    Dim GP As IGDIPlusParam
                
    Set GP = New GDIPlus
    
    GP.Initialize
    
    GP.EditType = GDIPlusConstants.GDIPlusCreateFromIPictureDisp
    Set GP.Pic = Pic
    
    GP.Run
                
    Set CreateObjectFromIPictureDisp = GP

End Function
Public Function CreateObject(ByVal Width As Long, ByVal Height As Long, Optional ByVal Color As Long = 0) As IGDIPlus
                    
    Dim GP As IGDIPlusParam
                
    Set GP = New GDIPlus
    
    GP.Initialize
    
    GP.EditType = GDIPlusConstants.GDIPlusCreateImage
    GP.Width = Width
    GP.Height = Height
    GP.Color = Color
    
    GP.Run
                
    Set CreateObject = GP

End Function
Private Property Let IGDIPlusParam_EditType(v As GDIPlusConstants)
    mEditType = v
End Property
Private Property Let IGDIPlusParam_FileName(v As String)
    mFileName = v
End Property
Private Property Let IGDIPlusParam_hIcon(v As LongPtr)
    mhIcon = v
End Property
Private Property Let IGDIPlusParam_Width(v As Long)
    mWidth = v
End Property
Private Property Let IGDIPlusParam_Height(v As Long)
    mHeight = v
End Property
Private Property Let IGDIPlusParam_Color(v As Long)
    mColor = v
End Property
Private Property Set IGDIPlusParam_Pic(v As IPictureDisp)
    Set mPic = v
End Property
Private Property Set IGDIPlusParam_Shape(v As Object)
    Set mShape = v
End Property

Private Sub IGDIPlusParam_Run()

        Dim result As Long

        Select Case mEditType
            Case GDIPlusCreateFromFile

                'BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF
                result = GdipCreateBitmapFromFile(StrPtr(mFileName), hGdiImage)

                Select Case result
                    Case Status.ok
                    Case Status.FileNotFound
                        MsgBox "file not found", vbExclamation, "GDIPlus"
                    Case Status.UnknownImageFormat
                        MsgBox "not support iamge format", vbExclamation, "GDIPlus"
                    Case Else
                        Err.Raise 512 + result, TypeName(Me) & ".Run", "GdipCreateBitmapFromFile error"
                End Select

            Case GDIPlusCreateFromIconHandle

                'hIcon → GdiPlusBitmap
                result = GdipCreateBitmapFromHICON(mhIcon, hGdiImage)
                If result <> Status.ok Then
                    Err.Raise 512 + result, TypeName(Me) & ".Run", "GdipCreateBitmapFromHICON error"
                End If

            Case GDIPlusCreateFromShape

                Dim hImg As LongPtr
                Dim hCopy As LongPtr
                Dim hPalette As LongPtr

                'クリップボードの保存
                If OpenClipboard(0&) = BOOL_FALSE Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Run", "OpenClipboard error"
                End If

                Call EmptyClipboard
                Call CloseClipboard

                '指定シェイプをビットマップでクリップボードに貼り付け
10              mShape.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

                Call CopyClipboardSleep

                If IsClipboardFormatAvailable(CF_BITMAP) = BOOL_FALSE Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Run", "IsClipboardFormatAvailable error"
                End If

                If OpenClipboard(0&) = BOOL_FALSE Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Run", "OpenClipboard error"
                End If

                hImg = GetClipboardData(CF_BITMAP)
                If hImg = 0 Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Run", "GetClipboardData error"
                End If

                hPalette = GetClipboardData(CF_PALETTE)

                hCopy = CopyImage(hImg, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
                If hCopy = 0 Then
                    Err.Raise 512 + 1, TypeName(Me) & ".Run", "CopyImage error"
                End If

                result = GdipCreateBitmapFromHBITMAP(hCopy, hPalette, hGdiImage)
                If result <> Status.ok Then
                    Err.Raise 512 + result, TypeName(Me) & ".Run", "GdipCreateBitmapFromHBITMAP error"
                End If

                Call EmptyClipboard
                Call CloseClipboard

            Case GDIPlusCreateFromIPictureDisp

                result = GdipCreateBitmapFromHBITMAP(mPic.handle, mPic.hpal, hGdiImage)
                If result <> Status.ok Then
                    Err.Raise 512 + result, TypeName(Me) & ".Run", "GdipCreateBitmapFromHBITMAP error"
                End If

            Case GDIPlusCreateImage

                '作業用hDC取得
                Dim hDc As LongPtr
                Dim hBitmap As LongPtr

                Dim lngWidth As Long
                Dim lngHeight As Long

                lngWidth = mWidth
                lngHeight = mHeight

                hDc = GetDC(Application.hWnd)

                hBitmap = CreateCompatibleBitmap(hDc, lngWidth, lngHeight)

                ReleaseDC Application.hWnd, hDc

                result = GdipCreateBitmapFromHBITMAP(hBitmap, 0&, hGdiImage)
                If result <> Status.ok Then
                    Err.Raise 512 + result, TypeName(Me) & ".Run", "GdipCreateBitmapFromHBITMAP error"
                End If

                Dim lngColor As Long

                lngColor = mColor

                IGDIPlus_DrawFillRectangle 0, 0, lngWidth, lngHeight, lngColor, 100

            Case Else
                Exit Sub
        End Select

End Sub
'--------------------------------------------------------------
' Initialize
'--------------------------------------------------------------
Private Sub IGDIPlusParam_Initialize()
   
    Dim uGdiInput As GdiplusStartupInput
    Dim result As Long
    
    uGdiInput.GdiplusVersion = 1&
    
    result = GdiplusStartup(hGdiPlus, uGdiInput)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".Class_Initialize", "GdiplusStartup error"
    End If
    
End Sub
'--------------------------------------------------------------
' Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    
    If hGdiImage <> 0 Then
        Call GdipDisposeImage(hGdiImage)
    End If
    
    If hGdiPlus <> 0 Then
        Call GdiplusShutdown(hGdiPlus)
    End If

End Sub
'alpha は %
Private Sub IGDIPlus_DrawFillRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)

    Dim hGraphics As LongPtr
    Dim hBrush As LongPtr
    Dim result As Long
    
    result = GdipGetImageGraphicsContext(hGdiImage, hGraphics)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipGetImageGraphicsContext error"
    End If
   
    result = GdipCreateSolidFill(ConvertColor(lColor, Alpha), hBrush)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipCreateSolidFill error"
    End If
    
    If Smoothing Then
        result = GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        If result <> Status.ok Then
            Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipSetSmoothingMode error"
        End If
    End If
    
    result = GdipFillRectangle(hGraphics, hBrush, X, Y, Width, Height)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipFillRectangle error"
    End If
        
    Call GdipDeleteBrush(hBrush)
    Call GdipDeleteGraphics(hGraphics)
    
End Sub
Private Sub IGDIPlus_DrawFillEllipse(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)

    Dim hGraphics As LongPtr
    Dim hBrush As LongPtr
    Dim result As Long
    
    result = GdipGetImageGraphicsContext(hGdiImage, hGraphics)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipGetImageGraphicsContext error"
    End If
   
    result = GdipCreateSolidFill(ConvertColor(lColor, Alpha), hBrush)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipCreateSolidFill error"
    End If
    
    If Smoothing Then
        result = GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        If result <> Status.ok Then
            Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipSetSmoothingMode error"
        End If
    End If
    
    result = GdipFillEllipse(hGraphics, hBrush, X, Y, Width, Height)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawFillRectangle", "GdipFillEllipse error"
    End If
        
    Call GdipDeleteBrush(hBrush)
    Call GdipDeleteGraphics(hGraphics)
    
End Sub
Private Function IGDIPlus_DrawRectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False) As Boolean

    Dim hGraphics As LongPtr
    Dim hPen As LongPtr
    Dim result As Long

    result = GdipGetImageGraphicsContext(hGdiImage, hGraphics)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipGetImageGraphicsContext error"
    End If
        
    result = GdipCreatePen1(ConvertColor(lColor, Alpha), 1, 2&, hPen)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipCreatePen1 error"
    End If
    
    If Smoothing Then
        result = GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        If result <> Status.ok Then
            Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipSetSmoothingMode error"
        End If
    End If
    
    result = GdipDrawRectangle(hGraphics, hPen, X, Y, Width, Height)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipDrawRectangle error"
    End If

    Call GdipDeletePen(hPen)
    Call GdipDeleteGraphics(hGraphics)
    
End Function
Private Sub IGDIPlus_DrawEllipse(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)

    Dim hGraphics As LongPtr
    Dim hPen As LongPtr
    Dim result As Long

    result = GdipGetImageGraphicsContext(hGdiImage, hGraphics)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipGetImageGraphicsContext error"
    End If
        
    result = GdipCreatePen1(ConvertColor(lColor, Alpha), 1, 2&, hPen)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipCreatePen1 error"
    End If
    
    If Smoothing Then
        result = GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        If result <> Status.ok Then
            Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipSetSmoothingMode error"
        End If
    End If
    
    result = GdipDrawEllipse(hGraphics, hPen, X, Y, Width, Height)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawRectangle", "GdipDrawEllipse error"
    End If

    Call GdipDeletePen(hPen)
    Call GdipDeleteGraphics(hGraphics)
    
End Sub
Private Sub IGDIPlus_DrawString(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal strFont As String, ByVal strBuf As String, ByVal Size As Long, Optional ByVal style As DrawStringConstants, Optional ByVal lColor As Long = 0, Optional ByVal Alpha As Long = 100, Optional ByVal Smoothing As Boolean = False)

    Dim hGraphics As LongPtr
    Dim hBrush As LongPtr
    Dim result As Long
    Dim hFormat As LongPtr
    
    result = GdipGetImageGraphicsContext(hGdiImage, hGraphics)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipGetImageGraphicsContext error"
    End If
   
    result = GdipCreateSolidFill(ConvertColor(lColor, Alpha), hBrush)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipCreateSolidFill error"
    End If
    
    If Smoothing Then
        result = GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias)
        If result <> Status.ok Then
            Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipSetSmoothingMode error"
        End If
    End If
    
    Dim hFontFamily As LongPtr
    
    result = GdipCreateFontFamilyFromName(StrPtr(strFont), 0, hFontFamily)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipCreateFontFamilyFromName error"
    End If
    
    Dim hFont As LongPtr
    Const Pixel As Long = 2
    
    GdipCreateFont hFontFamily, Size, style, Pixel, hFont
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipCreateFont error"
    End If

    Dim rect As RECTF
    
    rect.Top = Y
    rect.Left = X
    rect.Right = X + Width
    rect.Bottom = Y + Height
    

    
    'StringFormat
    result = GdipCreateStringFormat(0, LANG_NEUTRAL, hFormat)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipCreateStringFormat error"
    End If
    
    'StringFormat(Align)
    result = GdipSetStringFormatAlign(hFormat, StringAlignmentCenter)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipSetStringFormatAlign error"
    End If
    
    'StringFormat(vAlign)
    result = GdipSetStringFormatLineAlign(hFormat, StringAlignmentCenter)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipSetStringFormatLineAlign error"
    End If
    
    result = GdipDrawString(hGraphics, StrPtr(strBuf), -1, hFont, rect, hFormat, hBrush)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipDrawString error"
    End If
        
    result = GdipDeleteStringFormat(hFormat)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".DrawString", "GdipDeleteStringFormat error"
    End If
    
    Call GdipDeleteBrush(hBrush)
    Call GdipDeleteGraphics(hGraphics)
    
End Sub

'ピクチャに変換
Private Function IGDIPlus_ToPicture() As IPictureDisp
    
    Dim uPicInfo As PICTDESC
    Dim hBitmap As LongPtr
    Dim result As Long
    
    result = GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0&)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".ToPicture", "GdipCreateHBITMAPFromBitmap error"
    End If
    
    With uPicInfo
      .Size = LenB(uPicInfo)
      .Type = PICTYPE_BITMAP
      .hPic = hBitmap
      .hpal = 0&
    End With
        
    result = OleCreatePictureIndirect(uPicInfo, ConvGUID(IID_IPictureDisp), BOOL_TRUE, IGDIPlus_ToPicture)
    If result <> S_OK Then
        Err.Raise result, TypeName(Me) & ".ToPicture", "OleCreatePictureIndirect error"
    End If
    
End Function
Private Sub IGDIPlus_SaveAsJPEG(ByVal strFile As String, Optional ByVal Quality As Long = 85)

    Dim EncodParameters     As EncoderParameters
    Dim result As Long
    
    EncodParameters.Count = 1
    With EncodParameters.Parameter(0)
        
        .GUID = ConvGUID(CLSID_QUALITY)
        
        .NumberOfValues = 1
        ' 4=EncoderParameterValueTypeLong
        .Type = 4
        ' 圧縮品質
        .Value = VarPtr(Quality)
    End With
    
    ' JPG変換で保存
    result = GdipSaveImageToFile(hGdiImage, StrPtr(strFile), ConvGUID(CLSID_JPEG), VarPtr(EncodParameters))
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".SaveAsJPEG", "GdipSaveImageToFile error"
    End If

End Sub
Private Sub IGDIPlus_SaveAsPNG(ByVal strFile As String)

    Dim result As Long
    
    ' PNG変換で保存
    result = GdipSaveImageToFile(hGdiImage, StrPtr(strFile), ConvGUID(CLSID_PNG), 0&)
    If result <> Status.ok Then
        Err.Raise 512 + result, TypeName(Me) & ".SaveAsPNG", "GdipSaveImageToFile error"
    End If

End Sub
Private Function ConvertColor(Color As Long, Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
 
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    
    CopyMemory ConvertColor, BGRA(0), 4&

End Function
Private Function ConvGUID(ByVal strGUID As String) As GUID
    
    Call IIDFromString(StrPtr(strGUID), ConvGUID)

End Function
'------------------------------------------------------------------------------------------------------------------------
' CopyPictureがJavaアプリやクリップボードツールなどで失敗する対策
'------------------------------------------------------------------------------------------------------------------------
Private Sub CopyClipboardSleep()
    DoEvents
    Sleep Val(Registry.GetSetting("Option", "ClipboardSleep", 10))
End Sub
