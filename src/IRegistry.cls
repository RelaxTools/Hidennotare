VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'--------------------------------------------------------------
' レジストリ読み込み
'--------------------------------------------------------------
Public Function GetSetting(ByVal Section As String, ByVal key As String, Optional ByVal Default As String = "") As String
End Function
'--------------------------------------------------------------
' レジストリ書き込み
'--------------------------------------------------------------
Public Sub SaveSetting(ByVal Section As String, ByVal key As String, ByVal Value As String)
End Sub
'--------------------------------------------------------------
'レジストリ削除
'--------------------------------------------------------------
Public Sub DeleteSetting(ByVal Section As String, Optional ByVal key As String = "")
End Sub
'--------------------------------------------------------------
' レジストリのExport
' UTF-16 LE にてレジストリの内容を出力する。
'--------------------------------------------------------------
Public Sub Export()
End Sub


