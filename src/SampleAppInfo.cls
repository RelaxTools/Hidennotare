VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleAppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IAppInfo

Private mLogger As ILogger
Private mRegistry As IRegistry
Private mINIFile As IRegistry



Private Property Get IAppInfo_LogFolder() As String
    IAppInfo_LogFolder = FileIO.BuildPath(SystemInfo.GetAppFolder, ThisWorkbook.BuiltinDocumentProperties("Title").Value, "Log")
End Property

Private Property Get IAppInfo_LogHistory() As Long
    IAppInfo_LogHistory = 5
End Property

Private Property Get IAppInfo_LogLevel() As Hidennotare.LogLevel
    IAppInfo_LogLevel = Info
End Property

Private Property Get IAppInfo_LogType() As Hidennotare.LogType
    IAppInfo_LogType = Logfile
End Property

Private Property Get IAppInfo_AppName() As String
    IAppInfo_AppName = ThisWorkbook.BuiltinDocumentProperties("Title").Value
End Property

Private Property Get IAppInfo_INIFileName() As String
    IAppInfo_INIFileName = FileIO.BuildPath(SystemInfo.GetAppFolder, ThisWorkbook.BuiltinDocumentProperties("Title").Value, ThisWorkbook.BuiltinDocumentProperties("Title").Value & ".ini")
End Property

'Public Function Logger() As ILogger
'    If mLogger Is Nothing Then
'        Set mLogger = Hidennotare.Logger.NewInstance(ThisWorkbook)
'    End If
'    Set Logger = mLogger
'End Function

'Public Function Registry() As IRegistry
'    If mRegistry Is Nothing Then
'        Set mRegistry = Hidennotare.Registry.NewInstance(ThisWorkbook)
'    End If
'    Set Registry = mRegistry
'End Function
'Public Function INIFile() As IRegistry
'    If mINIFile Is Nothing Then
'        Set mINIFile = Hidennotare.INIFile.NewInstance(ThisWorkbook)
'    End If
'    Set INIFile = mINIFile
'End Function
