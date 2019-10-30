VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
'>### Graphics クラス
'>
'>**Remarks**
'>
'>PNG形式の読み込み書き込みが行えるクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-09-04|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Function ExtractIconExW Lib "Shell32" (ByVal lpszFile As LongPtr, ByVal nIconIndex As Long, phiconLarge As LongPtr, phiconSmall As LongPtr, ByVal nIcons As Long) As Long

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

'>---
'>#### LoadPicture(Graphics)
'>
'>**Syntax**
'>
'>```
'>Set var = Graphics.LoadPicture(filename)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Graphics|必須|Graphics|Graphicsクラスを指定します。|
'>|filename|必須|String|ロードする画像ファイル名を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|IPicture|VBAで使用可能なピクチャ形式|
'>
'>**Remarks**
'>
'>画像ファイルをVBAで表示できるPicture形式に変換します。<br>
'>対応フォーマット:BMP, GIF, JPEG, PNG, TIFF, Exif, WMF, and EMF
'>
'>**See also**
'>
'>* SavePicture
'>
Public Function LoadPicture(ByVal FileName As String) As IPictureDisp

    '存在チェック
    If Not FileIO.FileExists(FileName) Then
        Err.Raise 512 + 1, TypeName(Me) & "LoadPicture", "file not found"
    End If
    
    '形式チェック
    Select Case UCase(FileIO.GetExtensionName(FileName))
        Case "BMP"
        Case "GIF"
        Case "PNG"
        Case "JPG", "JPEG"
        Case "WMF", "EMF"
        Case "EXIF"
        Case "TIF", "TIFF"
        Case Else
            Err.Raise 512 + 1, TypeName(Me) & "LoadPicture", "Not Support Type "
    End Select

    Set LoadPicture = GDIPlus.CreateObjectFromFile(FileName).ToPicture

End Function
'>---
'>#### SavePicture(Graphics)
'>
'>**Syntax**
'>
'>```
'>Graphics.SavePicture(picture, filename)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Graphics|必須|Graphics|Graphicsクラスを指定します。|
'>|picture|必須|IPictureDisp|保存する画像のPictureプロパティを指定します。|
'>|filename|必須|String|保存する画像ファイル名を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>指定されたPictureをPNG形式で保存します。<br>
'>他、保存形式を使用したい場合にはGDIPlusクラスを利用すること。<br>
'>
'>**See also**
'>
'>* LoadPicture
'>
Public Sub SavePicture(Picture As IPictureDisp, FileName As String)

    If FileIO.FileExists(FileName) Then
    
        Call FileIO.DeleteFile(FileName)
    
    End If

    Call GDIPlus.CreateObjectFromIPictureDisp(Picture).SaveAsPNG(FileName)
    
End Sub

Public Function LoadIcon(ByVal FileName As String, Optional ByVal Size As IconConstants = IconConstants.IconSmall) As IPictureDisp

    Dim hIconLarge As LongPtr
    Dim hIconSmall As LongPtr
    
    Dim hIcon As LongPtr
    
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

    If hIcon = 0 Then
        Err.Raise 512 + 1, TypeName(Me) & ".LoadIcon", "hIcon = null "
    End If
    
    Set LoadIcon = GDIPlus.CreateObjectFromIconHandle(hIcon).ToPicture

    If hIcon <> 0 Then
        Call DestroyIcon(hIcon)
    End If

End Function
