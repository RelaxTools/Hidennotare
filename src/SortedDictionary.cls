VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedDictionary"
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
' SortedDictionary クラス
' 委譲 : Dictionary
'-----------------------------------------------------------------------------------------------------
' TODO : 本家と異なる点。For Each を使う場合には XX.Keys を省略できない。
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements IConstructor
Implements INewInstance

'VBAでは「継承」ができないため「委譲」する。
Private mDic As IDictionary
Private mCP As IComparer
Private mIndex As Long
Private mKeys As Variant
Private Sub Class_Initialize()
    Set mDic = New Dictionary
    mIndex = 0
End Sub
Private Sub Class_Terminate()
    Set mDic = Nothing
End Sub
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
Private Function IConstructor_Instancing(Args As Collection) As Object
    
    'コンパレータ
    If Args.Count = 1 Then
        Set mCP = Args(1)
    Else
        Exit Function
    End If
    
    Set IConstructor_Instancing = Me

End Function

Private Function IDictionary_Add(ByVal KeyValue As String, obj As Variant)
    mDic.Add KeyValue, obj
End Function
Private Function IDictionary_ContainsKey(ByVal KeyValue As String) As Boolean
    IDictionary_ContainsKey = mDic.ContainsKey(KeyValue)
End Function
'--------------------------------------------------------------
'  ソート済みキーリスト取得
'--------------------------------------------------------------
Private Property Get IDictionary_Keys() As Variant
    IDictionary_Keys = Sort(mDic.Keys)
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
Private Function IDictionary_ToString() As String
    IDictionary_ToString = mDic.ToString
End Function
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = Me
    mKeys = Sort(mDic.Keys)
End Property
'--------------------------------------------------------------
'  ICursor の Eof プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mIndex > mDic.Count - 1
End Property
'--------------------------------------------------------------
'  ICursor の Item プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    
    If mIndex > mDic.Count - 1 Then
        ICursor_Item = mKeys(mDic.Count - 1)
    Else
        ICursor_Item = mKeys(mIndex)
    End If

End Property
'--------------------------------------------------------------
'  ICursor の カーソル位置を初期化
'--------------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  ICursor の カーソル位置を次に移動
'--------------------------------------------------------------
Private Sub ICursor_MoveNext()

    mIndex = mIndex + 1
    If mIndex > mDic.Count Then
        mIndex = mDic.Count
    End If

End Sub
'--------------------------------------------------------------
'  キーのソート
'--------------------------------------------------------------
Private Function Sort(ByVal strKey As Variant) As Variant

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim wk As String
    
    If mCP Is Nothing Then
        Set mCP = New IComparer
    End If

    '要素数
    n = UBound(strKey) - LBound(strKey) + 1
    
    'ソート不要
    If n <= 1 Then
        GoTo e
    End If

    '挿入ソート
    For i = 1 To n - 1

        wk = strKey(i)
        
        Dim lngCmp As Long
        
        lngCmp = mCP.Compare(strKey(i - 1), wk)

        If lngCmp > 0 Then

            j = i

            Do

                strKey(j) = strKey(j - 1)

                j = j - 1

                If j = 0 Then
                    Exit Do
                End If

                lngCmp = mCP.Compare(strKey(j - 1), wk)
            
            Loop While lngCmp > 0
            
            strKey(j) = wk

        End If
    Next

e:
    Sort = strKey

End Function
'--------------------------------------------------------------
' インスタンス生成 I/F
'--------------------------------------------------------------
Private Function INewInstance_NewInstance() As Object
    Set INewInstance_NewInstance = New SortedDictionary
End Function
