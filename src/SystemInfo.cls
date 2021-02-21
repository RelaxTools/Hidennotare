VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemInfo"
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
'>### SystemInfo クラス
'>
'>**Remarks**
'>
'>SystemInfoクラス<br>
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
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Enum OSVersion
    OSVersionWindowsXP = 1
    OSVersionWindowsVista
    OSVersionWindows7
    OSVersionWindows8
    OSVersionWindows10
End Enum

Public Enum OSArchitecture
    OSArchitecture32bit = 1
    OSArchitecture64bit
End Enum
'>---
'>#### GetOSArchitecture(SystemInfo)
'>
'>**Syntax**
'>
'>```
'>ret = SystemInfo.GetOSArchitecture()
'>```
'>
'>**Parameters**
'>
'>なし
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|OSArchitecture|OSのアーキテクチャ(32bit/64bit)の定数を返却|
'>
'>
'>**Remarks**
'>
'>OSBit数取得<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* SystemInfo.GetOSVersion
'>
Public Function GetOSArchitecture() As OSArchitecture
    
    Dim v As Variant
    
    v = Split(Application.OperatingSystem)

    If Arrays.Length(v) >= 4 Then
        Select Case v(1)
            Case "(32-bit)"
                GetOSArchitecture = OSArchitecture32bit
            Case "(64-bit)"
                GetOSArchitecture = OSArchitecture64bit
            Case Else
        End Select
    Else
        Err.Raise 512 + 1, TypeName(Me) & ".GetOSArchitecture"
    End If

End Function
'>---
'>#### GetOSVersion(SystemInfo)
'>
'>**Syntax**
'>
'>```
'>ret = SystemInfo.GetOSVersion()
'>```
'>
'>**Parameters**
'>
'>なし
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|OSArchitecture|OSの種類/バージョン(WindowsXP,Vista,7,8,10)の定数を返却|
'>
'>
'>**Remarks**
'>
'>OSVersion取得<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* SystemInfo.GetOSArchitecture
'>
Public Function GetOSVersion() As OSVersion

    Dim v As Variant
    
    v = Split(Application.OperatingSystem)

    If Arrays.Length(v) >= 4 Then
        Select Case v(3)
            Case "5.01"
                GetOSVersion = OSVersionWindowsXP
            Case "6.00"
                GetOSVersion = OSVersionWindowsVista
            Case "6.01"
                GetOSVersion = OSVersionWindows7
            Case "6.02"
                GetOSVersion = OSVersionWindows8
            Case "10.00"
                GetOSVersion = OSVersionWindows10
            Case Else
        End Select
    Else
        Err.Raise 512 + 1, TypeName(Me) & ".GetOSVersion"
    End If

End Function
Public Function GetAppFolder() As String
    
    GetAppFolder = VBA.CreateObject("Wscript.Shell").SpecialFolders("AppData")

End Function
