VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderedDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
'>### OrderedDictionary クラス
'>
'>**Remarks**
'>
'>DictionaryのKeyの追加順を保証します。<br>
'>Dictionary クラスをコンポジット（移譲）したクラス。<br>
'>使用方法についてはDictionaryを参照。
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>- IDictionary I/F を使用する必要がある。
'>- Orderd なので追加順を維持するため削除のパフォーマンスは悪いので注意。
'>
'>**Implements**
'>
'>* ICursor
'>* IDictionary
'>* ICompatibleProperty
'>* INewInstance
'>
'>**Dependent**
'>
'>* Message
'>* StringBuilder
'>* ArrayList
'>* IDictionary
'>* INewInstance
'>* ICompatibleProperty
'>* StringUtils
'>* ICursor
'>
'>**Sample**
'>
'>```
'>Dim dic As IDictionary
'>Set dic = New OrderedDictionary
'>
'>dic.Add Key, Value
'>```
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements INewInstance
Implements IValueObject

'VBAでは「継承」ができないため「委譲」する。
Private mCol As Collection
Private mDic As Object
Private mCP As IComparer
Private mIndex As Long
Private mKeys As Variant
Private Sub Class_Initialize()
    
    Set mDic = VBA.CreateObject("Scripting.Dictionary")
    mDic.CompareMode = vbBinaryCompare
    
    mIndex = 0
    mKeys = Array()
End Sub
Private Sub Class_Terminate()
    Set mDic = Nothing
End Sub
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set obj = OrderdDictionary.NewInstance()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|OrderdDictionary|必須|OrderdDictionary|OrderdDictionaryを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|obj|IDictionary|新しいインスタンス(IDictionary)|
'>
'>**Remarks**
'>
'> OrderdDictionary を生成する場合
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
'>---
Public Function NewInstance() As IDictionary

    Set NewInstance = New OrderedDictionary

    Call Class_Terminate

End Function





'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IDictionary_NewEnum() As stdole.IUnknown

    Set mCol = Arrays.ToCollection(mKeys)
    Set IDictionary_NewEnum = mCol.[_NewEnum]

End Function
Private Function Self() As IDictionary
    Set Self = Me
End Function
Private Sub IDictionary_Add(KeyValue As Variant, obj As Variant)
    
    mDic.Add KeyValue, obj
    
    '配列にキーーを保存
    ReDim Preserve mKeys(0 To mDic.Count - 1)
    mKeys(mDic.Count - 1) = KeyValue

End Sub
Private Function IDictionary_ContainsKey(KeyValue As Variant) As Boolean
    IDictionary_ContainsKey = mDic.Exists(KeyValue)
End Function
'--------------------------------------------------------------
'  追加順キーリスト取得
'--------------------------------------------------------------
Private Function IDictionary_Keys() As Variant
    
    IDictionary_Keys = mKeys

End Function
Private Property Get IDictionary_Count() As Long
    IDictionary_Count = mDic.Count
End Property
Private Property Get IDictionary_Item(KeyValue As Variant) As Variant
    
    Objects.SetWith(IDictionary_Item) = mDic.Item(KeyValue)
    
End Property
Private Property Let IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
Private Property Set IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
Private Property Let IDictionary_Key(OldKey As Variant, NewKey As Variant)
    
    mDic.key(OldKey) = NewKey
    
    Dim i As Long
    For i = LBound(mKeys) To UBound(mKeys)
        If mKeys(i) = OldKey Then
            mKeys(i) = NewKey
            Exit Property
        End If
    Next
    
End Property
Private Sub IDictionary_Remove(KeyValue As Variant)
    
    mDic.Remove KeyValue
    
    If mDic.Count = 0 Then
        mKeys = Array()
    Else
    
        Dim lngCnt As Long
        Dim i As Long
        Dim j As Long
        Dim NewKeys As Variant
        
        NewKeys = Array()
        
        lngCnt = mDic.Count - 1
        
        ReDim NewKeys(0 To lngCnt)
        
        j = 0
        For i = LBound(mKeys) To UBound(mKeys)
            If mKeys(i) <> KeyValue Then
                NewKeys(j) = mKeys(i)
                j = j + 1
            End If
        Next
    
        mKeys = NewKeys
    
    End If

End Sub
Private Sub IDictionary_Clear()
    mDic.Clear
    mKeys = Array()
End Sub
Private Function IDictionary_ToString() As String
    IDictionary_ToString = ToString
End Function
Private Function ToString() As String

    Dim Item  As Variant
    Dim KeyValue As Variant
    
    Dim v  As Variant
    
    Dim SB As IStringBuilder
    Set SB = StringBuilder.NewInstance()
    
    For Each KeyValue In mDic.Keys
    
        Objects.SetWith(v) = mDic.Item(KeyValue)
    
        Select Case True
            Case Objects.InstanceOfIDictionary(v)
                
                SB.Append """" & KeyValue & """:" & Objects.AsIDictionary(v).ToString
            
            Case Objects.InstanceOfIList(v)
                
                SB.Append """" & KeyValue & """:" & Objects.AsIList(v).ToString
            
            Case Objects.InstanceOfIValueObject(v)
                
                SB.Append """" & KeyValue & """:" & Objects.AsIValueObject(v).ToString
            
            Case IsEmpty(v)
                
                SB.Append """" & KeyValue & """:" & "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        SB.Append """" & KeyValue & """:" & LCase(CStr(v))
        
                    Case vbString
                    
                        SB.Append """" & KeyValue & """:""" & StringUtils.Escape(v) & """"
                    
#If Win64 Then
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbLongLong
#Else
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
#End If
        
                        SB.Append """" & KeyValue & """:" & CStr(v)
        
                    Case Else
        
                        'エラー
                        Err.Raise 512 + 1, TypeName(Me) & ".ToString", "not cast type (" & TypeName(v) & ")"
        
                End Select

        End Select
    
    Next

    ToString = SB.ToString(", ", "{", "}")

End Function
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = Me
End Property
Private Property Set IDictionary_Comparer(RHS As IComparer)
    Err.Raise 512 + 1, TypeName(Me) & ".Comparer", "Not Support Property"
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
' インスタンス生成 I/F
'--------------------------------------------------------------
Private Function INewInstance_NewInstance() As Object

    Dim d As IDictionary
    
    Set d = New OrderedDictionary

    Set INewInstance_NewInstance = d

End Function
'--------------------------------------------------------------
' IValueObject I/F
'--------------------------------------------------------------
Private Property Let IValueObject_Fields(key As Variant, RHS As Variant)
    Self.Add key, RHS
End Property
Private Property Get IValueObject_Fields(key As Variant) As Variant
    Objects.SetWith(IValueObject_Fields) = Self.Item(key)
End Property
Private Property Get IValueObject_Count() As Long
    IValueObject_Count = Self.Count
End Property
Private Property Get IValueObject_FieldNames() As Variant
    IValueObject_FieldNames = Self.Keys
End Property
Private Function IValueObject_ToString() As String
    IValueObject_ToString = Self.ToString
End Function
