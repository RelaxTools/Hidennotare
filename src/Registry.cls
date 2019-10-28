VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registry"
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
'>### Registry クラス
'>
'>**Remarks**
'>
'>レジストリ管理<br>
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
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------
'レジストリ読み込み
'--------------------------------------------------------------
Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
    GetSetting = VBA.GetSetting(Core.GetTitleName, Section, Key, Default)
End Function
'--------------------------------------------------------------
'レジストリ書き込み
'--------------------------------------------------------------
Public Sub SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Value As String)
    VBA.SaveSetting Core.GetTitleName, Section, Key, Value
End Sub
'--------------------------------------------------------------
'レジストリ削除
'--------------------------------------------------------------
Public Sub DeleteSetting(ByVal Section As String, Optional ByVal Key As String = "")
    On Error Resume Next
    If Len(Key) = 0 Then
        VBA.DeleteSetting Core.GetTitleName, Section
    Else
        VBA.DeleteSetting Core.GetTitleName, Section, Key
    End If
End Sub
'--------------------------------------------------------------
' レジストリのExport
' UTF-16 LE にてレジストリの内容を出力する。
'--------------------------------------------------------------
Public Sub Export()

    Dim strDat As String
    Const C_FF As Byte = &HFF
    Const C_FE As Byte = &HFE
    Dim FileName As Variant
    Dim strReg As String
    Dim Key As String
    
    Dim Reg, Locator, Service, SubKey, RegName, RegType
    Dim i As Long, j As Long, buf As String, RegData As String
    
    Dim FP As Integer
    
    FileName = Application.GetSaveAsFilename(InitialFileName:=Core.GetTitleName & ".reg", fileFilter:="登録ファイル,*.reg")
    If FileName = False Then
        Exit Sub
    End If
    
    On Error GoTo err_Handle

    strReg = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\" & Core.GetTitleName

    Set Locator = VBA.CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer(vbNullString, "root\default")
    Set Reg = Service.Get("StdRegProv")
    
    Const HKEY_CURRENT_USER = &H80000001
    
    Const ROOT = "HKEY_CURRENT_USER\"
    Key = "Software\VB and VBA Program Settings\" & Core.GetTitleName
    
    Reg.EnumKey HKEY_CURRENT_USER, Key, SubKey
    
    FP = FreeFile()
    Open FileName For Output As FP
    Close FP
    
    FP = FreeFile()
    Open FileName For Binary As FP
    
    Dim strBuf() As Byte
    
    Put FP, , C_FF
    Put FP, , C_FE
    
    strBuf = "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf
    Put FP, , strBuf
    
    strBuf = "[" & ROOT & Key & "]" & vbCrLf
    Put FP, , strBuf
    
    For i = 0 To UBound(SubKey)
        
        Reg.EnumValues HKEY_CURRENT_USER, Key & "\" & SubKey(i), RegName, RegType
            
        strBuf = vbCrLf & "[" & ROOT & Key & "\" & SubKey(i) & "]" & vbCrLf
        Put FP, , strBuf
        
        For j = 0 To UBound(RegName)
        
            Select Case RegType(j)
                Case 1
                    Reg.GetStringValue HKEY_CURRENT_USER, Key & "\" & SubKey(i), RegName(j), RegData
                Case Else
                    Reg.GetMultiStringValue HKEY_CURRENT_USER, Key & "\" & SubKey(i), RegName(j), RegData
                
            End Select
        
            strDat = Replace(RegData, "\", "\\")
            strDat = Replace(strDat, """", "\""")
            
            strBuf = """" & RegName(j) & """=""" & strDat & """" & vbCrLf
            
            Put FP, , strBuf
        
        Next j
        
    Next i
    strBuf = vbCrLf
    Put FP, , strBuf
    Close FP
    
    Set Reg = Nothing
    Set Service = Nothing
    Set Locator = Nothing
    
    Message.Information "登録ファイルを保存しました。\n移行先で登録ファイルを実行するとレジストリに反映されます。"
    Exit Sub

err_Handle:
    Message.Information "登録ファイルの保存に失敗しました。"
End Sub

