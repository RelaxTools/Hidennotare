VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Implements IRegistry
Implements IAppNameParam

Private mAppName As String

Private Declare PtrSafe Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpsectionName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpDefault As LongPtr, ByVal lpReturnedString As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long

Private Declare PtrSafe Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long



Public Function NewInstance(ai As IAppInfo) As IRegistry

    Dim p As IAppNameParam
    
    Set p = New IniFile

    p.Name = ai.INIFileName

    Set NewInstance = p

End Function
Private Property Let IAppNameParam_Name(RHS As String)
    mAppName = RHS
End Property

Private Function IRegistry_GetSetting(ByVal Section As String, ByVal key As String, Optional ByVal Default As String = "") As String

    Dim lngRet As Long
    Dim ret As String
    Dim size As Long
    
    ret = String$(512, vbNullChar)
    size = Len(ret)
    
    lngRet = GetPrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Default), StrPtr(ret), size, StrPtr(mAppName))
    
    
    IRegistry_GetSetting = Left$(ret, InStr(ret, vbNullChar) - 1)

End Function

Private Sub IRegistry_SaveSetting(ByVal Section As String, ByVal key As String, ByVal Value As String)

    Dim lngRet As Long

    lngRet = WritePrivateProfileStringW(StrPtr(Section), StrPtr(key), StrPtr(Value), StrPtr(mAppName))

End Sub

Private Sub IRegistry_DeleteSetting(ByVal Section As String, Optional ByVal key As String = "")
    
    Dim lngRet As Long
    lngRet = WritePrivateProfileStringW(StrPtr(Section), StrPtr(key), 0, StrPtr(mAppName))

End Sub

Private Sub IRegistry_Export()


End Sub
