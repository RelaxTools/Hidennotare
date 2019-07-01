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
' Power Query M Language List Class(Collection Wrapper)
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
Private Sub IList_Insert(Idx As Long, v As Variant)
    mCol.Insert Idx, v
End Sub
Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = mCol.NewEnum
End Function
Private Property Get IList_Count() As Long
    IList_Count = mCol.Count
End Property
Private Property Get IList_Item(Idx As Long) As Variant
    If VBA.IsObject(mCol.Item(Idx)) Then
        Set IList_Item = mCol.Item(Idx)
    Else
        IList_Item = mCol.Item(Idx)
    End If
End Property
Private Sub IList_RemoveAt(Idx As Long)
    mCol.RemoveAt Idx
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
' MList‚ð•¶Žš—ñ‚É•ÏŠ·
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

    IList_ToString = "{" & SB.ToJoin(", ") & "}"

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
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    ICursor_PreviousItem = mCol.GetCursor.PreviousItem
End Property
Private Sub ICursor_MoveFirst()
    mCol.GetCursor.MoveFirst
End Sub
Private Sub ICursor_MoveNext()
    mCol.GetCursor.MoveNext
End Sub
