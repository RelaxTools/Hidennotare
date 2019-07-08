VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [MLangBuilder] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/MFunctionCreater
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
' Power Query M Language Record Class(Dictionary Wrapper)
' 委譲 : Dictionary
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary

'VBAでは「継承」ができないため「委譲」する。
Private mDic As IDictionary
Private Sub Class_Initialize()
    Set mDic = New Dictionary
End Sub
Private Sub Class_Terminate()
    Set mDic = Nothing
End Sub
Private Function IDictionary_Add(ByVal KeyValue As String, obj As Variant)
    mDic.Add KeyValue, obj
End Function
Private Function IDictionary_ContainsKey(ByVal KeyValue As String) As Boolean
    IDictionary_ContainsKey = mDic.ContainsKey(KeyValue)
End Function
Private Property Get IDictionary_Keys() As Variant
    IDictionary_Keys = mDic.Keys
End Property
Private Property Get IDictionary_Count() As Long
    IDictionary_Count = mDic.Count
End Property
Private Property Get IDictionary_Item(ByVal KeyValue As String) As Variant
    If VBA.IsObject(mDic.Item(KeyValue)) Then
        Set IDictionary_Item = mDic.Item(KeyValue)
    Else
        IDictionary_Item = mDic.Item(KeyValue)
    End If
End Property
Private Property Let IDictionary_Item(ByVal KeyValue As String, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
Private Property Set IDictionary_Item(ByVal KeyValue As String, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
Private Property Let IDictionary_Key(ByVal OldKey As String, NewKey As Variant)
    mDic.Key(OldKey) = NewKey
End Property
Private Sub IDictionary_Remove(ByVal KeyValue As String)
    mDic.Remove KeyValue
End Sub
Private Sub IDictionary_Clear()
    mDic.Clear
End Sub
'--------------------------------------------------------------
'  シリアライズ化(MRecord)
'--------------------------------------------------------------
Private Function IDictionary_ToString() As String

    Dim Item  As Variant
    Dim KeyValue As Variant
    
    Dim v  As Variant
    
    Dim SB As StringBuilder
    Set SB = New StringBuilder
    
    For Each KeyValue In mDic.Keys
    
        If VBA.IsObject(mDic.Item(KeyValue)) Then
            Set v = mDic.Item(KeyValue)
        Else
            v = mDic.Item(KeyValue)
        End If
    
        Select Case True
            
            Case IsObject(v)
                
                SB.Append KeyValue & "=" & v.ToString
            
            Case Else

                SB.Append KeyValue & "=" & v

        End Select
    
    Next

    IDictionary_ToString = "[" & SB.ToJoin(", ") & "]"

End Function
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = mDic.GetCursor
End Property
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mDic.GetCursor.Eof
End Property
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    ICursor_Item = mDic.GetCursor.Item
End Property
Private Sub ICursor_MoveFirst()
    mDic.GetCursor.MoveFirst
End Sub
Private Sub ICursor_MoveNext()
    mDic.GetCursor.MoveNext
End Sub


