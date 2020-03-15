VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValueObject

Public Name As String
Public Age As Long
Public Address As String

Public mFields As Variant


Private Sub Class_Initialize()
    mFields = Array("Name", "Age", "Address")
End Sub

Private Property Let IValueObject_Fields(key As Variant, RHS As Variant)
    CallByName Me, key, VbLet, RHS
End Property

Private Property Get IValueObject_Fields(key As Variant) As Variant
    IValueObject_Fields = CallByName(Me, key, VbGet)
End Property

Private Property Get IValueObject_Count() As Long
    IValueObject_Count = Arrays.Length(mFields)
End Property

Private Property Get IValueObject_FieldNames() As Variant
    IValueObject_FieldNames = Array("Name", "Age", "Address")
End Property

Private Function IValueObject_ToString() As String
    
    Dim SB As IStringBuilder
    
    With StringBuilder.NewInstance
    
        .Append Name, True
        .Append CStr(Age)
        .Append Address, True
    
        IValueObject_ToString = .ToString(",", "{", "}")
    
    End With
    
End Function
