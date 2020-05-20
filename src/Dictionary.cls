VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2.5
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hidennotare
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
'>### Dictionary クラス
'>
'>**Remarks**
'>
'>Scripting.Dictionary のコンポジット（移譲）クラス。
'>
'>- IDictionary I/F を使用する必要がある。
'>
'>**Scripting.Dictionaryとの比較**
'>
'>|No.|Property/Method|Scripting.Dictionary|Dictionary(本クラス)|
'>|---|---|---|---|
'>|1|Add|〇|〇|
'>|2|Exists|〇|ContainsKey|
'>|3|Items|〇|なし|
'>|4|Keys|〇|〇|
'>|5|Remove|〇|〇|
'>|6|RemoveAll|〇|Clear|
'>|7|CompareMode|〇|なし(BinaryCompare固定)|
'>|8|Count|〇|〇|
'>|9|Item|〇|〇|
'>|10|Key|〇|〇|
'>|11|ToString|×|〇|
'>
'>**Sample**
'>
'>```
'>Dim dic As IDictionary
'>Set dic = Dictionary.NewInstance
'>
'>dic.Add Key, Value
'>```
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
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>|2020-05-20|Ver.1.0.1|NewEnumの実装をコレクションを作成するのではなく、CallByNameで取得する方法に修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements INewInstance
Implements IValueObject

'Private mCol As Collection
Private mDic As Object
Private mIndex As Long
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set obj = Dictionary.NewInstance()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Dictionary|必須|Dictionary|Dictionaryを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|obj|IDictionary|新しいインスタンス(IDictionary)|
'>
'>**Remarks**
'>
'> Dictionary を生成する場合
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

    Set NewInstance = New Dictionary
    Call Class_Terminate

End Function
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    
    Set mDic = VBA.CreateObject("Scripting.Dictionary")
    mDic.CompareMode = vbBinaryCompare
    mIndex = 0

End Sub
'--------------------------------------------------------------
'  Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    
    Set mDic = Nothing

End Sub
'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IDictionary_NewEnum() As stdole.IUnknown

'    Set mCol = Arrays.ToCollection(mDic.Keys)
'    Set IDictionary_NewEnum = mCol.[_NewEnum]
    'Dictionary の列挙子、オブジェクトブラウザには表示されないが『_NewEnum』Twitter @nukie_53
    Set IDictionary_NewEnum = CallByName(mDic, "_NewEnum", VbMethod)

End Function
'>---
'>#### Add(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Add(Key, obj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|キーを指定|
'>|obj|必須|Variant|追加するオブジェクトを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary に項目を追加する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Sub IDictionary_Add(KeyValue As Variant, obj As Variant)
    
    mDic.Add KeyValue, obj

End Sub
'>---
'>#### ContainsKey(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>bool = expression.ContainsKey(Key)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|キー項目|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|bool|Boolean|キーが存在する場合:True/キーが存在しない場合:False|
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary のキーの存在チェック
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Function IDictionary_ContainsKey(KeyValue As Variant) As Boolean
    
    IDictionary_ContainsKey = mDic.Exists(KeyValue)

End Function
'>---
'>#### Keys(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Keys()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|キー配列を返却|
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary のキー情報を返却
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Function IDictionary_Keys() As Variant
    
    IDictionary_Keys = mDic.Keys

End Function
'>---
'>#### Count(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Count
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Long|Dictionaryの要素数を返却|
'>
'>**Remarks**
'>
'>IDictionary I/F のプロパティです。<br>
'>Dictionary 要素数を返却
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Property Get IDictionary_Count() As Long
    
    IDictionary_Count = mDic.Count

End Property
'>---
'>#### Item(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Item(idx)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|取得する項目のキーを指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|Dictionary の指定位置の項目|
'>
'>**Remarks**
'>
'>IDictionary I/F のGetプロパティです。<br>
'>Dictionary の指定位置の項目を取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Property Get IDictionary_Item(KeyValue As Variant) As Variant
    
    Objects.SetWith(IDictionary_Item) = mDic.Item(KeyValue)

End Property
'>---
'>#### Item(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>Let expression.Item(Key) = value
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|取得する項目のキーを指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|value|Variant|新しい設定項目|
'>
'>**Remarks**
'>
'>IDictionary I/F のLetプロパティです。<br>
'>Dictionary の指定キーの項目を設定
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Property Let IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    
    mDic.Item(KeyValue) = NewItem

End Property
'>---
'>#### Item(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>Set expression.Item(Key) = value
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|取得する項目のキーを指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|value|Variant|新しい設定項目|
'>
'>**Remarks**
'>
'>IDictionary I/F のSetプロパティです。<br>
'>Dictionary の指定キーの項目を設定
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Property Set IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    
    Set mDic.Item(KeyValue) = NewItem

End Property
'--------------------------------------------------------------
'  キー設定
'--------------------------------------------------------------
Private Property Let IDictionary_Key(OldKey As Variant, NewKey As Variant)
    
    mDic.key(OldKey) = NewKey

End Property
'>---
'>#### Remove(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Remove(Key)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>|Key|必須|Variant|削除する項目のキーを指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary の指定位置の項目を削除
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Sub IDictionary_Remove(KeyValue As Variant)
    
    mDic.Remove KeyValue

End Sub
'>---
'>#### Clear(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Clear()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary の項目を全削除
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Sub IDictionary_Clear()
    
    mDic.RemoveAll
    mIndex = 0

End Sub
'>---
'>#### ToString(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToString()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|JSON文字列を返却|
'>
'>**Remarks**
'>
'>IDictionary I/F のメソッドです。<br>
'>Dictionary をJSONに変換
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Private Function IDictionary_ToString() As String
    
    Dim Item  As Variant
    Dim KeyValue As Variant
    
    Dim v  As Variant
    
    With StringBuilder.NewInstance
    
        For Each KeyValue In mDic.Keys
        
            Objects.SetWith(v) = mDic.Item(KeyValue)
        
            Select Case True
                Case Objects.InstanceOfIDictionary(v)
                    
                    .Append """" & KeyValue & """:" & Objects.AsIDictionary(v).ToString
                
                Case Objects.InstanceOfIList(v)
                    
                    .Append """" & KeyValue & """:" & Objects.AsIList(v).ToString
                
                Case Objects.InstanceOfIValueObject(v)
                    
                    .Append """" & KeyValue & """:" & Objects.AsIValueObject(v).ToString
                
                Case IsEmpty(v)
                    
                    .Append """" & KeyValue & """:" & "null"
                
                Case Else
                    Select Case VarType(v)
                        
                        Case vbBoolean
            
                            .Append """" & KeyValue & """:" & LCase(CStr(v))
            
                        Case vbString
                        
                            .Append """" & KeyValue & """:""" & StringUtils.Escape(v) & """"
                        
#If Win64 Then
                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbLongLong
#Else
                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
#End If
            
                            .Append """" & KeyValue & """:" & CStr(v)
            
                        Case Else
            
                            'エラー
                            Err.Raise 512 + 1, TypeName(Me) & ".ToString", "not cast type (" & TypeName(v) & ")"
            
                    End Select
    
            End Select
        
        Next
    
        IDictionary_ToString = .ToString(", ", "{", "}")
    
    End With

End Function
'>---
'>#### GetCursor(IDictionary I/F)
'>
'>**Syntax**
'>
'>```
'>Set var = expression.GetCursor()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IDictionary|IDictionary I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|ICursor|ICursor I/F を返却|
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ICursor インターフェースを取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>* ICursor
'>
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = Me
End Property
Private Property Set IDictionary_Comparer(RHS As IComparer)
    
    Err.Raise 512 + 1, TypeName(Me) & ".Comparer", "Not Support Property"

End Property
'>---
'>#### Eof(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Eof()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Boolean|カーソルが終了:True / カーソルが終了していない:False|
'>
'>**Remarks**
'>
'>ICursor I/F のプロパティです。<br>
'>ICursor の Eof プロパティを取得。カーソルの終了を検出します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Property Get ICursor_Eof() As Boolean
    
    ICursor_Eof = mIndex > mDic.Count - 1

End Property
'>---
'>#### Item(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.Item()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|カーソル位置の内容を取得します。|
'>
'>**Remarks**
'>
'>ICursor I/F のプロパティです。<br>
'>ICursor の Item プロパティを取得。現在行の内容を取得します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    
    Dim v As Variant
    
    v = IDictionary_Keys
    
    If mIndex > mDic.Count - 1 Then
        ICursor_Item = v(mDic.Count - 1)
    Else
        ICursor_Item = v(mIndex)
    End If

End Property
'>---
'>#### MoveFirst(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>expression.MoveFirst()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ICursor I/F のメソッドです。<br>
'>カーソル位置の初期化を行います。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Sub ICursor_MoveFirst()
    
    mIndex = 0

End Sub
'>---
'>#### MoveNext(ICursor I/F)
'>
'>**Syntax**
'>
'>```
'>expression.MoveNext()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|ICursor I/F を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ICursor I/F のメソッドです。<br>
'>カーソル位置を次に移動します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Private Sub ICursor_MoveNext()

    mIndex = mIndex + 1
    If mIndex > mDic.Count Then
        mIndex = mDic.Count
    End If

End Sub
'>---
'>#### NewInstance(INewInstance I/F)
'>
'>**Syntax**
'>
'>```
'>Set var = expression.NewInstance()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICursor|INewInstance I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Object|新しいDictionaryのインスタンス(IDictionary I/F)|
'>
'>**Remarks**
'>
'>INewInstance I/F のメソッドです。<br>
'>新しいインスタンス(IDictionary I/F)を生成して返却します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>* INewInstance
'>
Private Function INewInstance_NewInstance() As Object
    
    Set INewInstance_NewInstance = Me.NewInstance

End Function
'>---
'>#### SetProperty(ICompatibleProperty I/F)
'>
'>**Syntax**
'>
'>```
'>expression.SetProperty(property, value)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICompatibleProperty|ICompatibleProperty I/F を指定します。|
'>|property|必須|Variant|プロパティ名を指定|
'>|value|必須|Variant|値を設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Object|新しいDictionaryのインスタンス(IDictionary I/F)|
'>
'>**Remarks**
'>
'>SetProperty I/F のメソッドです。<br>
'>指定プロパティキーに値を設定する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>* ICompatibleProperty
'>
'Private Sub ICompatibleProperty_SetProperty(prop As Variant, val As Variant)
'
'    If Self.ContainsKey(prop) Then
'        Objects.SetWith(Self.Item(prop)) = val
'    Else
'        Self.Add prop, val
'    End If
'
'End Sub
'>---
'>#### GetProperty(ICompatibleProperty I/F)
'>
'>**Syntax**
'>
'>```
'>Set var = expression.GetProperty()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|ICompatibleProperty|ICompatibleProperty I/F を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Object|新しいDictionaryのインスタンス(IDictionary I/F)|
'>
'>**Remarks**
'>
'>GetProperty I/F のメソッドです。<br>
'>指定プロパティキーの値を取得する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>* ICompatibleProperty
'>
'Private Function ICompatibleProperty_GetProperty(prop As Variant) As Variant
'
'    Objects.SetWith(ICompatibleProperty_GetProperty) = Self.Item(prop)
'
'End Function

Private Function Self() As IDictionary

    Set Self = Me

End Function
Private Property Let IValueObject_Fields(key As Variant, RHS As Variant)
    
    If Self.ContainsKey(key) Then
        Objects.SetWith(Self.Item(key)) = RHS
    Else
        Self.Add key, RHS
    End If

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
    IValueObject_ToString = Self.ToString()
End Function
