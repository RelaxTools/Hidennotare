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
' [Hecatoncheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hecatoncheir
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
'   なし
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成
'-----------------------------------------------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------
'レジストリ読み込み
'--------------------------------------------------------------
Public Function GetSetting(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
    GetSetting = VBA.GetSetting(Title, Section, Key, Default)
End Function
'--------------------------------------------------------------
'レジストリ書き込み
'--------------------------------------------------------------
Public Sub SaveSetting(ByVal Section As String, ByVal Key As String, ByVal Value As String)
    VBA.SaveSetting Title, Section, Key, Value
End Sub
'--------------------------------------------------------------
'レジストリ削除
'--------------------------------------------------------------
Public Sub DeleteSetting(ByVal Section As String, Optional ByVal Key As String = "")
    If Len(Key) = 0 Then
        VBA.DeleteSetting Title, Section
    Else
        VBA.DeleteSetting Title, Section, Key
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
    
    Dim fp As Integer
    
    FileName = Application.GetSaveAsFilename(InitialFileName:=Title & ".reg", fileFilter:="登録ファイル,*.reg")
    If FileName = False Then
        Exit Sub
    End If
    
    On Error GoTo err_Handle

    strReg = "HKEY_CURRENT_USER\Software\VB and VBA Program Settings\" & Title

    Set Locator = CreateObject("WbemScripting.SWbemLocator")
    Set Service = Locator.ConnectServer(vbNullString, "root\default")
    Set Reg = Service.Get("StdRegProv")
    
    Const HKEY_CURRENT_USER = &H80000001
    
    Const ROOT = "HKEY_CURRENT_USER\"
    Key = "Software\VB and VBA Program Settings\" & Title
    
    Reg.EnumKey HKEY_CURRENT_USER, Key, SubKey
    
    fp = FreeFile()
    Open FileName For Output As fp
    Close fp
    
    fp = FreeFile()
    Open FileName For Binary As fp
    
    Dim strBuf() As Byte
    
    Put fp, , C_FF
    Put fp, , C_FE
    
    strBuf = "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf
    Put fp, , strBuf
    
    strBuf = "[" & ROOT & Key & "]" & vbCrLf
    Put fp, , strBuf
    
    For i = 0 To UBound(SubKey)
        
        Reg.EnumValues HKEY_CURRENT_USER, Key & "\" & SubKey(i), RegName, RegType
            
        strBuf = vbCrLf & "[" & ROOT & Key & "\" & SubKey(i) & "]" & vbCrLf
        Put fp, , strBuf
        
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
            
            Put fp, , strBuf
        
        Next j
        
    Next i
    strBuf = vbCrLf
    Put fp, , strBuf
    Close fp
    
    Set Reg = Nothing
    Set Service = Nothing
    Set Locator = Nothing
    
    Message.Information "登録ファイルを保存しました。\n移行先で登録ファイルを実行するとレジストリに反映されます。"
    Exit Sub

err_Handle:
    Message.Information "登録ファイルの保存に失敗しました。"
End Sub
'--------------------------------------------------------------
'アプリケーション名取得
'--------------------------------------------------------------
Private Property Get Title() As String
    Title = ThisWorkbook.BuiltinDocumentProperties("Title").Value
End Property
