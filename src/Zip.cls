VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Zip"
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
'>### Zip クラス
'>
'>**Remarks**
'>
'>Zip クラス<br>
'>Zip ファイルの操作が可能(PowerShellを利用。.NetFramework 4.5以上必須)<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>- None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### CompressArchive(Zip)
'>
'>**Syntax**
'>
'>```
'>Zip.CompressArchive(col, zipfile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Zip|必須|Zip|Zipオブジェクトを指定します。|
'>|col|必須|Collection|圧縮するファイルをCollectionで指定|
'>|zipfile|必須|String|作成するZipファイルを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>Zip圧縮<br>
'>指定されたファイルを圧縮して、Zipファイルを作成します。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Zip.ExpandArchive
'>
Public Sub CompressArchive(ByVal ary As Variant, ByVal strDest As String)

    Dim SB As StringBuilder
    Dim PH As StringBuilder
    Dim v As Variant
    
    If Not strDest Like "*.zip" Then
        Err.Raise 512 + 1, TypeName(Me) & ".CompressArchive", "support zip only"
    End If
    
    Dim strParent As String
    strParent = FileIO.GetParentFolderName(strDest)
    If Not FileIO.FolderExists(strParent) Then
        Err.Raise 512 + 1, TypeName(Me) & ".CompressArchive", "not exists folder " & strParent
    End If
    
    If Not IsArray(ary) Then
        ary = Array(ary)
    End If
    
    For Each v In ary
        If Not FileIO.FolderExists(v) And Not FileIO.FileExists(v) Then
            Err.Raise 512 + 1, TypeName(Me) & ".CompressArchive", "not exists file/folder " & v
        End If
    Next
    
    Set SB = New StringBuilder
    
    'コマンド
    SB.Append "Compress-Archive"
    
    '圧縮するファイル
    SB.Append "-Path"
    Set PH = New StringBuilder
    For Each v In ary
    
        PH.Append v, True
    
    Next
    SB.Append PH.ToString(",")
    
    '出力パス
    SB.Append "-DestinationPath"
    SB.Append strDest, True

    SB.Append "-Force"
    
    'PowerShell を実行する
    Process.ExecPowerShell SB.ToString(" ")

End Sub
'>---
'>#### ExpandArchive(Zip)
'>
'>**Syntax**
'>
'>```
'>Zip.ExpandArchive(zipfile, path)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Zip|必須|Zip|Zipオブジェクトを指定します。|
'>|zipfile|必須|Collection|Zipファイル指定|
'>|path|必須|String|展開するフォルダを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>Zip展開<br>
'>Zipファイルを指定されたフォルダに展開します。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* Zip.CompressArchive
'>
Public Sub ExpandArchive(ByVal strArchive As String, ByVal strPath As String)

    If Not strArchive Like "*.zip" Then
        Err.Raise 512 + 1, TypeName(Me) & ".ExpandArchive", "support zip only"
    End If
    
    If Not FileIO.FileExists(strArchive) Then
        Err.Raise 512 + 1, TypeName(Me) & ".ExpandArchive", "zip file not exists"
    End If
    
    If Not FileIO.FolderExists(strPath) Then
        Err.Raise 512 + 1, TypeName(Me) & ".ExpandArchive", "folder not exists"
    End If
    
    Dim SB As StringBuilder
    
    Set SB = New StringBuilder
    
    'コマンド
    SB.Append "Expand-Archive"
    
    'ZIPファイル
    SB.Append "-LiteralPath"
    SB.Append strArchive, True
    
    '出力パス
    SB.Append "-DestinationPath"
    SB.Append strPath, True

    SB.Append "-Force"
    
    'PowerShell を実行する
    Process.ExecPowerShell SB.ToString(" ")

End Sub
