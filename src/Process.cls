VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Process"
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
'>### Process クラス
'>
'>**Remarks**
'>
'>プロセス管理<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* ExcelHelper
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-16|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub SleepLocal Lib "kernel32" Alias "Sleep" (ByVal ms As LongPtr)
Private RefCount As Long
'--------------------------------------------------------------
' マルチプロセス実行用（マクロ）
'--------------------------------------------------------------
Public Sub MultiProsess(ByVal strMacro As String)
    
    Err.Clear
    
    On Error Resume Next
    With VBA.CreateObject("Excel.Application")
        .Workbooks.Open ThisWorkbook.FullName
        .Run ExcelUtils.BuildOnAction(strMacro)
    End With
    
    If Err.Number <> 0 Then
        MsgBox "Multi Process起動に失敗しました。", vbCritical
    End If
    
End Sub
'--------------------------------------------------------------
' 非同期実行用（マクロ）
'--------------------------------------------------------------
Function UnsyncRun(ByVal strMacro As String, Optional ByVal lngSec As Long = 0) As Long
    Application.OnTime DateAdd("s", lngSec, Now), strMacro
End Function
'--------------------------------------------------------------
'　標準出力取得実行
'--------------------------------------------------------------
Function Exec(ByVal strCommand As String) As String

    Dim WSH As Object
    Dim wExec As Object
    
    Set WSH = VBA.CreateObject("WScript.Shell")
    
    Set wExec = WSH.Exec(strCommand)
    Do While wExec.Status = 0
        DoEvents
    Loop
    
    Exec = wExec.StdOut.ReadAll
    
    Set wExec = Nothing
    Set WSH = Nothing

End Function
'--------------------------------------------------------------
'　アドレス文字列からオブジェクトに変換
'--------------------------------------------------------------
Public Function GetObjectFromAddress(ByVal strAddress As String) As Object

    Dim obj As Object

    Dim p As LongPtr
    p = CLngPtr(strAddress)
  
    CopyMemory obj, p, LenB(p)
    
    Set GetObjectFromAddress = obj

End Function
'--------------------------------------------------------------
'　Shellコマンド
'--------------------------------------------------------------
Sub Shell(ByVal strCommand As String, ParamArray p())

    Dim ret As Variant
    Dim strBuf As String
    Dim v As Variant

    strBuf = """" & strCommand & """"
    
    If UBound(p) >= 0 Then

        For Each v In p
            strBuf = strBuf & " """ & v & """"
        Next

    End If
    
    On Error Resume Next
    ret = VBA.Shell(strBuf, vbNormalFocus)
    If ret <> 0 Then
        AppActivate ret
    End If

End Sub
'--------------------------------------------------------------
'　OpenFile コマンド(ファイル実行)
'--------------------------------------------------------------
Sub OpenFile(ByVal strFile As String, Optional ByVal flg As Boolean = False)

    On Error GoTo e

    With VBA.CreateObject("WScript.Shell")
        .Run """" & strFile & """", 1, flg
    End With

    Exit Sub
e:
    MsgBox "ファイルを開けませんでした。" & strFile, vbCritical
End Sub
'--------------------------------------------------------------
'　同期バッチ実行
'--------------------------------------------------------------
Function ExecBatch(ByVal strFile As String) As Long

    On Error GoTo e

    With VBA.CreateObject("WScript.Shell")
        ExecBatch = .Run(strFile, 0, True)
    End With

    Exit Function
e:
    MsgBox "ファイルを開けませんでした。" & strFile, vbCritical
End Function
'--------------------------------------------------------------
'　PowerShell 実行
'--------------------------------------------------------------
Function ExecPowerShell(ByVal strCommand As String) As Long
    
    Dim strFile As String
    Dim SB As IStringBuilder
    
    strFile = FileIO.BuildPath(FileIO.TempFolder, FileIO.GetTempName & ".ps1")
    
    'テキスト出力
    TextWriter.NewInstance(strFile).WriteData strCommand
    
    On Error GoTo e
    
    'PowerShell を実行
    Set SB = StringBuilder.NewInstance
    
    SB.Append "powershell"
    SB.Append "-ExecutionPolicy"
    SB.Append "RemoteSigned"
    SB.Append "-File"
    SB.Append strFile, True

    Process.ExecBatch SB.ToString(" ")
    
e:
    FileIO.DeleteFile strFile

End Function
'--------------------------------------------------------------
'　Sleep
'--------------------------------------------------------------
Function Sleep(ByVal lngtime As Long)
    SleepLocal lngtime
End Function
