VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
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
'>### Logger クラス
'>
'>**Remarks**
'>
'>ログ出力クラス<br>
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
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ILogger
Implements ILoggerParam

#If Win64 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As LongPtr
#Else
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#End If
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

'スタック
Private colStack As New Collection

'ファイル番号
Private FP As Integer


Private mFSO As Object

Private mLevel As LogLevel
Private mLogType As LogType
Private mlngHistory As Long
Private mLevelStr As Variant
Private mName As String

Private mOpen As Boolean

Private mFolder As String

Public Function NewInstance(ai As IAppInfo) As ILogger

    Dim p As ILoggerParam
    
    Set p = New Logger

    p.Name = ai.AppName
    p.History = ai.LogHistory
    p.Level = ai.LogLevel
    p.LogType = ai.LogType
    p.Folder = ai.LogFolder
    
    p.Init

    Set NewInstance = p

End Function
Private Property Let ILoggerParam_Folder(RHS As String)
    mFolder = RHS
End Property
Private Property Let ILoggerParam_History(RHS As Long)
    mlngHistory = RHS
End Property

Private Property Let ILoggerParam_Level(RHS As LogLevel)
    mLevel = RHS
End Property

Private Property Let ILoggerParam_LogType(RHS As LogType)
    mLogType = RHS
End Property

Private Property Let ILoggerParam_Name(RHS As String)
    mName = RHS
End Property

Private Sub ILoggerParam_Init()

    Set mFSO = VBA.CreateObject("Scripting.FileSystemObject")

    FP = 225

    mLevelStr = Array("Trace", "Info", "Warn", "Fatal", "None")

    If mLogType And LogType.Logfile Then
        Call OpenLog
    End If

End Sub
'--------------------------------------------------------------
'　Class_Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()

End Sub
'--------------------------------------------------------------
'　Class_Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()

    On Error Resume Next
    
    If mOpen Then
        Close FP
        Set mFSO = Nothing
    End If

End Sub
'--------------------------------------------------------------
' アプリケーション名
'--------------------------------------------------------------
'Private Property Let ILogger_AppName(v As String)
'    mAppName = v
'End Property
'>---
'>#### LogTrace(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogTrace(message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>トレースログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogTrace(m As String, ParamArray p())
    ReportLog LogLevel.Trace, PlaceHolder(m, p)
End Sub
'>---
'>#### LogInfo(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogInfo(message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>情報ログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogInfo(m As String, ParamArray p())
    ReportLog LogLevel.Info, PlaceHolder(m, p)
End Sub
'>---
'>#### LogWarn(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogWarn(message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ワーニングログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogWarn(m As String, ParamArray p())
    ReportLog LogLevel.Warn, PlaceHolder(m, p)
End Sub
'>---
'>#### LogFatal(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogFatal(message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>エラーログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogFatal(m As String, ParamArray p())

    Dim lngErr As Long
    Dim strDesc As String

    lngErr = Err.Number
    strDesc = Err.Description

    ReportLog LogLevel.Fatal, "Fatal Error (" & PlaceHolder(m, p) & ")"
    ILogger_LogFlash

    Err.Number = lngErr
    Err.Description = strDesc

End Sub
'>---
'>#### LogBegin(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogBegin(level, message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|level|必須|LogLevel|出力するログレベルを指定|
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>開始ログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogBegin(lvl As LogLevel, m As String, ParamArray p())

    If colStack Is Nothing Then
        Set colStack = New Collection
    End If

    ReportLog lvl, "[" & colStack.Count + 1 & "]" & PlaceHolder(m, p) & ",BEGIN"

    colStack.Add GetTickCount

End Sub
'>---
'>#### LogFinish(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogFinish(level, message, [param])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|level|必須|LogLevel|出力するログレベルを指定|
'>|message|必須|String|ログに出力するメッセージを指定|
'>|param|必須|ParamArray|パラメータ文字列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>終了ログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogFinish(lvl As LogLevel, m As String, ParamArray p())

    Dim t As LongPtr

    If colStack Is Nothing Then
        t = 0
    Else
        If colStack.Count = 0 Then
            t = 0
        Else
            t = colStack.Item(colStack.Count)
            colStack.Remove colStack.Count
        End If
    End If

    If t = 0 Then
        ReportLog lvl, PlaceHolder(m, p) & ",FINISH,[?]ms"
    Else
        t = GetTickCount - t
        ReportLog lvl, "[" & colStack.Count + 1 & "]" & PlaceHolder(m, p) & ",FINISH,[" & t & "]ms"
    End If
    
    ILogger_LogFlash

End Sub
'>---
'>#### LogFlash(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.LogFlash()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>終了ログ出力<br>
'>
'>**See also**
'>
'>* None
'>
Private Sub ILogger_LogFlash()
    Call OpenLog
End Sub
'>---
'>#### Level(Logger)
'>
'>**Syntax**
'>
'>```
'>Logger.Level = value
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|value|必須|LogLevel|出力するログレベルを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ログレベルの設定<br>
'>
'>**See also**
'>
'>* None
'>
Private Property Let ILogger_Level(ByVal l As LogLevel)
    mLevel = l
End Property
'ログ出力
Private Sub ReportLog(ByVal l As LogLevel, ByVal strMsg As String)

    Dim strLog As String
    Dim lngRetry As Long

    On Error GoTo e

    If l >= mLevel Then
        strLog = mName & "," & Format$(Now, "yyyy-mm-dd,hh:nn:ss") & "." & getMSec & "," & mLevelStr(l) & "," & strMsg
        If mLogType And LogType.DebugWindow Then
            Debug.Print strLog
        End If
        If mLogType And LogType.Logfile Then
            Print #FP, strLog
        End If
    End If

    Exit Sub
e:
    If Err.Number = 52 Then 'ファイル名または番号が不正です。
        Call OpenLog
        lngRetry = lngRetry + 1
        Sleep 500
        If lngRetry > 10 Then
            MsgBox "Logger Failed(" & Err.Description & ")", vbCritical
        Else
            Resume
        End If
    End If
End Sub
'時間の取得(ms)
Private Function getMSec() As String

    Dim dblTimer As Double

    dblTimer = CDbl(Timer)
    getMSec = Format$((dblTimer - Fix(dblTimer)) * 1000, "000")

End Function
'ログのオープン
Private Sub OpenLog()

    Dim strFile As String
    Dim strLogFolder As String
    Dim i As Long

    On Error Resume Next
    Close FP

    On Error GoTo e

    'ログフォルダ作成
    strLogFolder = AppDataFolder

    Dim col As Collection
    Dim objFolder As Object
    Dim objFile As Object

    Set col = New Collection

    'ログフォルダを Collection に追加
    Set objFolder = mFSO.GetFolder(strLogFolder)
    For Each objFile In objFolder.files

        For i = 1 To col.Count
            If col(i) >= objFile.Name Then
                col.Add objFile.Name, , i
                GoTo pass
            End If
        Next
        col.Add objFile.Name
pass:
    Next

    Dim v As Variant

    'フォルダ名の昇順に処理する
    i = col.Count
    For Each v In col

        If i > mlngHistory Then
            Kill strLogFolder & "\" & v
        End If
        i = i - 1
    Next

    Set col = Nothing
    Set objFolder = Nothing

    strFile = mFSO.BuildPath(strLogFolder, Format(Now, "yyyymmdd") & ".log")
    Open strFile For Append As FP
    
    mOpen = True

    Exit Sub
e:
    Debug.Print "Logger File Open Error " & Err.Description
End Sub
'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Private Property Get AppDataFolder() As String

    On Error Resume Next

    AppDataFolder = ""

    With mFSO

        If .FolderExists(mFolder) Then
        Else
            CreateFolder mFolder
        End If

        AppDataFolder = .BuildPath(mFolder, "\")

    End With

End Property
'--------------------------------------------------------------
'  フォルダの作成（再帰）
'--------------------------------------------------------------
Private Sub CreateFolder(ByVal strPath As String)

    With mFSO

        If .FolderExists(strPath) Then
            Exit Sub
        End If

        Call CreateFolder(.GetParentFolderName(strPath))

        .CreateFolder strPath

    End With

End Sub
'----------------------------------------------------------------------------------
'プレースホルダ変換
'----------------------------------------------------------------------------------
Private Function PlaceHolder(ByVal strMsg As String, ByVal p As Variant) As String

    strMsg = StringUtils.Unescape(strMsg)

    If UBound(p) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(p)
            strMsg = Replace(strMsg, "{" & CStr(i) & "}", p(i))
        Next
    End If

    PlaceHolder = strMsg

End Function



