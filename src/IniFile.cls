VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFile"
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
'>### IniFile クラス
'>
'>**Remarks**
'>
'>INIファイルを操作するクラス<br>
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2021-02-23|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IRegistry
Implements IAppNameParam

Private mINIFile As String

Private Declare PtrSafe Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpsectionName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpDefault As LongPtr, ByVal lpReturnedString As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set objct = IniFile.NewInstance(appInfo)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|IniFile|必須|InFile|InFileクラスを指定|
'>|appInfo|必須|IAppinfo|IAppInfo I/F のインスタンスを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|objct|IRegistry|新しいインスタンス(IRegistry)|
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function NewInstance(ai As IAppInfo) As IRegistry

    Dim p As IAppNameParam
    
    Set p = New IniFile

    p.Name = ai.IniFileName

    Set NewInstance = p

End Function
Private Property Let IAppNameParam_Name(RHS As String)
    mINIFile = RHS
End Property
'>---
'>#### GetSetting(IRegistry I/F)
'>
'>**Syntax**
'>
'>```
'>ret = expression.GetSetting(Section, Key, [Default])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IRegistry|IRegistry I/F のインスタンス|
'>|Section|必須|String|値を取得するカテゴリ|
'>|Key|必須|String|値を取得するキー|
'>|Defalut|任意|String|キーが存在しなかった場合のデフォルトを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---|---
'>|ret|String|INIファイルに保存されている設定値|
'>
'>**Remarks**
'>
'> IRegistry I/F のメソッド<br>
'> INIファイル読み込み
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IRegistry
'>* Registry
'>
Private Function IRegistry_GetSetting(ByVal Section As String, ByVal key As String, Optional ByVal Default As String = "") As String

    Dim lngRet As Long
    Dim ret As String
    Dim size As Long
    
    ret = String$(512, vbNullChar)
    size = Len(ret)
    
    lngRet = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(ret), size, StrPtr(mINIFile))
    IRegistry_GetSetting = Left$(ret, InStr(ret, vbNullChar) - 1)

End Function
'>---
'>#### SaveSetting(IRegistry I/F)
'>
'>**Syntax**
'>
'>```
'>expression.SaveSetting(Section, Key, Value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IRegistry|IRegistry I/F のインスタンス|
'>|Section|必須|String|値を設定するカテゴリ|
'>|Key|必須|String|値を設定するキー|
'>|Value|必須|String|設定する値を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'> IRegistry I/F のメソッド<br>
'> INIファイル書き込み
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IRegistry
'>* Registry
'>
Private Sub IRegistry_SaveSetting(ByVal Section As String, ByVal key As String, ByVal Value As String)

    Dim lngRet As Long

    lngRet = WritePrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Value), StrPtr(mINIFile))

End Sub
'>---
'>#### DeleteSetting(IRegistry I/F)
'>
'>**Syntax**
'>
'>```
'>expression.DeleteSetting(Section, [Key])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IRegistry|IRegistry I/F のインスタンス|
'>|Section|必須|String|削除するカテゴリ|
'>|Key|任意|String|削除するキー|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'> IRegistry I/F のメソッド<br>
'> INIファイルのセクションまたはキーの削除
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IRegistry
'>* Registry
'>
Private Sub IRegistry_DeleteSetting(ByVal Section As String, Optional ByVal key As String = "")
    
    Dim lngRet As Long
    
    If key = "" Then
        lngRet = WritePrivateProfileStringW(StrPtr(Section), 0, 0, StrPtr(mINIFile))
    Else
        lngRet = WritePrivateProfileStringW(StrPtr(Section), StrPtr(key), 0, StrPtr(mINIFile))
    End If
End Sub
'>---
'>#### Export(IRegistry I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Export()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IRegistry|IRegistry I/F のインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'> IRegistry I/F のメソッド<br>
'> レジストリのExport
'> UTF-16 LE にてレジストリの内容を出力する
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IRegistry
'>* Registry
'>
Private Sub IRegistry_Export()

    Dim FileName As Variant
    
    FileName = Application.GetSaveAsFilename(InitialFileName:=mINIFile & ".ini", fileFilter:="INIファイル,*.ini")
    If FileName = False Then
        Exit Sub
    End If
    
    FileIO.CopyFile mINIFile, FileName
    
    Message.Information "INIファイルを保存しました。"
    Exit Sub

err_Handle:
    Message.Information "INIファイルの保存に失敗しました。"

End Sub
