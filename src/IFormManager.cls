VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get IsCancel() As Boolean

End Property

Public Property Get IsRunning() As Boolean

End Property

Public Sub StartRunning()

End Sub

Public Sub StopRunning()

End Sub

Public Sub DispGuidance(ByVal strValue As String)

End Sub

Public Sub StartGauge(ByVal lngValue As Long)

End Sub

Public Sub DisplayGauge(ByVal lngValue As Long)
    
End Sub

Public Sub doCancel()

End Sub

Public Sub SetHistory(c As ComboBox, ByVal strHistory As String)

End Sub

Public Function CreateHistory(c As ComboBox, Optional ByVal Max As Long = 10) As String

End Function


