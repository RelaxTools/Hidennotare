VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MList"
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
'>### MList クラス
'>
'>**Remarks**
'>
'>MList<br>
'>Power Query M Language List Class<br>
'>M言語をVBAで組み立てる際に便利なクラス<br>
'>
'* M 言語の List が作成できるクラス
'* IList I/F を使用する必要がある。
'* ToString すると M Language List を出力する。
'>
'>**Example**
'>
'>```
'>Dim lst As IList
'>Set lst = New MList
'>
'>lst.Add Value
'>```
'>
'>**Dependent**
'>
'>* IList
'>* ArrayList
'>* ICursor
'>* IConstructor
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor
Implements IList

Private mCol As IList
Private v() As Variant
Private mIndex As Long
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mCol = New ArrayList
End Sub
'--------------------------------------------------------------
'  Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Dim IC As IConstructor
    Set IC = mCol
    Set IConstructor_Instancing = IC.Instancing(Args)
    
End Function
Private Sub IList_Clear()
    mCol.Clear
End Sub
Private Sub IList_Add(obj As Variant)
    mCol.Add obj
End Sub
Private Sub IList_Insert(idx As Long, v As Variant)
    mCol.Insert idx, v
End Sub
Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = mCol.NewEnum
End Function
Private Property Get IList_Count() As Long
    IList_Count = mCol.Count
End Property
Private Property Get IList_Item(idx As Long) As Variant
    If VBA.IsObject(mCol.Item(idx)) Then
        Set IList_Item = mCol.Item(idx)
    Else
        IList_Item = mCol.Item(idx)
    End If
End Property
Private Sub IList_RemoveAt(idx As Long)
    mCol.RemoveAt idx
End Sub
Private Sub IList_Sort(Optional ByVal CP As IComparer = Nothing)
    mCol.Sort CP
End Sub
Private Function IList_ToArray() As Variant
    IList_ToArray = mCol.ToArray
End Function
Private Function IList_ToCollection() As Collection
    Set IList_ToCollection = mCol.ToCollection
End Function
'--------------------------------------------------------------
' MListを文字列に変換
'--------------------------------------------------------------
Private Function IList_ToString() As String
    
    Dim v As Variant
    Dim SB As StringBuilder
    Set SB = New StringBuilder
    
    For Each v In mCol
    
        Select Case True
            Case IsObject(v)

                SB.Append v.ToString

            Case Else

                SB.Append v

        End Select
    
    Next

    IList_ToString = SB.ToString(", ", "{", "}")

End Function
Private Property Get IList_GetCursor() As ICursor
    Set IList_GetCursor = mCol.GetCursor
End Property
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mCol.GetCursor.Eof
End Property
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    ICursor_Item = mCol.GetCursor.Item
End Property
Private Sub ICursor_MoveFirst()
    mCol.GetCursor.MoveFirst
End Sub
Private Sub ICursor_MoveNext()
    mCol.GetCursor.MoveNext
End Sub
