VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
' 依存モジュール
'   なし
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As LongPtr)

Private RefCount As Long
'--------------------------------------------------------------
' マルチプロセス実行用（マクロ）
'--------------------------------------------------------------
Public Sub MultiProsess(ByVal strMacro As String)
    
    Err.Clear
    
    On Error Resume Next
    With CreateObject("Excel.Application")
        .Workbooks.Open ThisWorkbook.FullName
        .Run ExcelHelper.BuildOnAction(strMacro)
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
    
    Set WSH = CreateObject("WScript.Shell")
    
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

    Dim Ret As Variant
    Dim strBuf As String
    Dim v As Variant

    strBuf = """" & strCommand & """"
    
    If UBound(p) >= 0 Then

        For Each v In p
            strBuf = strBuf & " """ & v & """"
        Next

    End If
    
    On Error Resume Next
    Ret = VBA.Shell(strBuf, vbNormalFocus)
    If Ret <> 0 Then
        AppActivate Ret
    End If

End Sub
'--------------------------------------------------------------
'　OpenFile コマンド(ファイル実行)
'--------------------------------------------------------------
Sub OpenFile(ByVal strFile As String, Optional ByVal flg As Boolean = False)

    On Error GoTo e

    With CreateObject("WScript.Shell")
        .Run """" & strFile & """", 1, flg
    End With

    Exit Sub
e:
    MsgBox "ファイルを開けませんでした。" & strFile, vbCritical
End Sub
'--------------------------------------------------------------
'　同期バッチ実行
'--------------------------------------------------------------
Sub ExecBatch(ByVal strFile As String)

    On Error GoTo e

    With CreateObject("WScript.Shell")
        .Run strFile, 0, True
    End With

    Exit Sub
e:
    MsgBox "ファイルを開けませんでした。" & strFile, vbCritical
End Sub
