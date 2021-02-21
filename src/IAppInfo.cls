VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get AppName() As String

End Property
Public Property Get LogFolder() As String

End Property
Public Property Get LogLevel() As LogLevel

End Property

Public Property Get LogType() As LogType

End Property

Public Property Get LogHistory() As Long

End Property

Public Property Get INIFileName() As String

End Property
