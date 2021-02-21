VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ログタイプ
Public Enum LogType
    DebugWindow = 1
    Logfile
    All
End Enum

'ログレベル
Public Enum LogLevel
    Trace = 0
    Info = 1
    Warn = 2
    Fatal = 3
    None = 4
End Enum

Public Sub LogTrace(m As String, ParamArray p())

End Sub
Public Sub LogInfo(m As String, ParamArray p())

End Sub
Public Sub LogWarn(m As String, ParamArray p())

End Sub
Public Sub LogFatal(m As String, ParamArray p())

End Sub
Public Sub LogBegin(lvl As LogLevel, m As String, ParamArray p())

End Sub
Public Sub LogFinish(lvl As LogLevel, m As String, ParamArray p())
    
End Sub
Public Sub LogFlash()

End Sub
