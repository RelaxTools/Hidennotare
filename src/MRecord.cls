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
'>### MRecord クラス
'>
'>**Remarks**
'>
'>MRecord<br>
'>Power Query M Language Record Class<br>
'>M言語をVBAで組み立てる際に便利なクラス<br>
'>
'* M 言語の Record が作成できるクラス
'* IDictionary I/F を使用する必要がある。
'* ToString すると M Language Record を出力する。
'>
'>**Example**
'>
'>```
'>Dim dic As IDictionary
'>Set dic = New MRecord
'>
'>dic.Add Key, Value
'>```
'>
'>**Dependent**
'>
'>* Dictionary
'>* IDictionary
'>* ICursor
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary

'VBAでは「継承」ができないため「委譲」する。
Private mCol As Collection
Private mDic As Object
Private Sub Class_Initialize()
    Set mDic = VBA.CreateObject("Scripting.Dictionary")
    mDic.CompareMode = vbBinaryCompare
End Sub
Private Sub Class_Terminate()
    Set mDic = Nothing
End Sub


'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IDictionary_NewEnum() As stdole.IUnknown

    Set mCol = Core.ArrayToCollection(mDic.Keys)
    Set IDictionary_NewEnum = mCol.[_NewEnum]

End Function
Private Sub IDictionary_Add(KeyValue As Variant, obj As Variant)
    mDic.Add KeyValue, obj
End Sub
Private Function IDictionary_ContainsKey(KeyValue As Variant) As Boolean
    IDictionary_ContainsKey = mDic.Exists(KeyValue)
End Function
Private Function IDictionary_Keys() As Variant
    IDictionary_Keys = mDic.Keys
End Function
Private Property Get IDictionary_Count() As Long
    IDictionary_Count = mDic.Count
End Property
Private Property Get IDictionary_Item(KeyValue As Variant) As Variant
    Core.SetVar(IDictionary_Item) = mDic.Item(KeyValue)
End Property
Private Property Let IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
Private Property Set IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
Private Property Let IDictionary_Key(OldKey As Variant, NewKey As Variant)
    mDic.Key(OldKey) = NewKey
End Property
Private Sub IDictionary_Remove(KeyValue As Variant)
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
    
    Dim SB As IStringBuilder
    Set SB = StringBuilder.CreateObject
    
    For Each KeyValue In mDic.Keys
    
        Core.SetVar(v) = mDic.Item(KeyValue)
    
        Select Case True
            
            Case IsObject(v)
                
                SB.Append KeyValue & "=" & v.ToString
            
            Case Else

                SB.Append KeyValue & "=" & v

        End Select
    
    Next

    IDictionary_ToString = SB.ToString(", ", "[", "]")

End Function
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = mDic.GetCursor
End Property
Private Property Set IDictionary_Comparer(RHS As IComparer)
    Err.Raise 512 + 1, TypeName(Me) & ".Comparer", "Not Support Property"
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



