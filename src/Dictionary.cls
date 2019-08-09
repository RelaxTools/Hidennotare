VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v1
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
'>Set dic = New Dictionary
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
'>* Convert
'>* ICursor
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements INewInstance
Implements ICompatibleProperty

Private mCol As Collection
Private mDic As Object
Private mIndex As Long
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mDic = CreateObject("Scripting.Dictionary")
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

    Set mCol = Constructor(New Collection, mDic.Keys)
    Set IDictionary_NewEnum = mCol.[_NewEnum]

End Function
'>---
'>#### GetIDictionary(Dictionary)
'>
'>**Syntax**
'>
'>```
'> Set obj = Dictionary.GetIDictionary()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|obj|IDictionary|現在のインスタンス(IDictionary I/F)|
'>
'>**Remarks**
'>
'>IDictionary I/F を取得する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Public Function GetIDictionary() As IDictionary
    Set GetIDictionary = Me
End Function
'>---
'>#### ParseFromListbox(Dictionary)
'>
'>**Syntax**
'>
'>```
'>Set ret = Dictionary.ParseFromListbox(objList, Key, [obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Object|Parseする ListBox or ComboBox を指定|
'>|Key|必須|String|行全体のキーとなる項目名を指定(DBMSの主キーに相当)|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Dictionary|新しいインスタンス(Dictionary)|
'>
'>
'>**Remarks**
'>
'>ListBox/ComboBox パース<br>
'>ListBox/ComboBoxからDictionary(レコードはDictionary)を生成する。<br>
'>Tagにレコードを生成するためのフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    'これはコードではなくプロパティウィンドウで設定しても良い
'>    lstFavorite.Tag = "No,Name,Address"
'>
'>    Dim dic As IDictionary
'>    Set dic = Dictionary.ParseFromListbox(lstFavorite, "Key", New FavoriteBean)
'>
'>```
'>
'>**See also**
'>
'>* Dictionary.CopyToListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListbox(objList As Object, KeyName As String, Optional obj As Object = Nothing) As IDictionary

    Dim ary As IDictionary
    Dim objRec As Object
    Dim Fields As Variant
    Dim i As Long
    Dim j As Long
    Dim Key As Variant
    
    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseFromListbox", "Item name is not set in tag"
    End If
    
    '配列に変換
    Fields = Split(objList.Tag, ",")
    
    '1次元目はArrayList固定
    Set ary = New Dictionary
        
    For i = 0 To objList.ListCount - 1
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = Core.GetNewInstance(New Dictionary)
            Case Else
                Set objRec = Core.GetNewInstance(obj)
        End Select
        
        For j = 0 To objList.ColumnCount - 1
        
            '設定されているフィールドのみコピーする。
            If Fields(j) <> "" Then
                Core.CastICompatibleProperty(objRec).SetProperty Fields(j), objList.List(i, j)
                
                If KeyName = Fields(j) Then
                    Key = objList.List(i, j)
                End If
            
            End If
        Next
            
        ary.Add Key, objRec
    
    Next

    Set ParseFromListbox = ary

End Function
'>---
'>#### CopyToListbox(Dictionary)
'>
'>**Syntax**
'>
'>```
'>Dictionary.CopyToListbox objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Object|copyする ListBox or ComboBox を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ListBox/ComboBox へコピー<br>
'>ArrayListの内容をListBox/ComboBoxにコピーする。<br>
'>ListBox/ComboBoxのTagにDictionaryのレコードを対応づけるフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    'これはコードではなくプロパティウィンドウで設定しても良い
'>    lstFavorite.Tag = "No,Name,Address"
'>
'>    Dictionary.CopyToListbox lstFavorite
'>
'>```
'>
'>**See also**
'>
'>* Dictionary.ParseFromListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToListbox(objList As Object)

    Dim ary As IDictionary
    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant
    Dim Key As Variant

    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".CopyToListbox", "Item name is not set in tag"
    End If

    '配列に変換
    Fields = Split(objList.Tag, ",")

    '1次元目はArrayList固定
    Set ary = Me

    'リストボックスのクリア
    objList.Clear

    Dim i As Long
    Dim j As Long
    i = 0
    For Each Key In ary.Keys
    
        Set objRec = ary.Item(Key)
    
        j = 0
        objList.AddItem
        For Each v In Fields

            '設定されているフィールドのみコピーする。
            If v <> "" Then
                objList.List(i, j) = Core.CastICompatibleProperty(objRec).GetProperty(v)
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### ParseFromListObject(Dictionary)
'>
'>**Syntax**
'>
'>```
'>Set ret = Dictionary.ParseFromListObject(objList, KeyField, [obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Listobject|Parseする Listobject を指定|
'>|KeyField|必須|String|行全体のキーとなる項目名を指定(DBMSの主キーに相当)|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Dictionary|新しいインスタンス(Dictionary)|
'>
'>**Remarks**
'>
'>ListObject パース<br>
'>ListobjectからDictionary(レコードはDictionary)を生成する。<br>
'>ListObject(テーブル)の列ヘッダの情報が必須<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>Sub Dictionary_ParseFromListObject_Sample()
'>
'>    Dim lst As IDictionary
'>    Dim dic As IDictionary
'>    Dim Key As Variant
'>
'>    Set lst = Dictionary.ParseFromListObject(ActiveSheet.ListObjects(1), "Key")
'>
'>    For Each dic In lst
'>
'>        For Each Key In dic.Keys
'>
'>            Debug.Print dic.Item(Key)
'>
'>        Next
'>
'>    Next
'>
'>    Dim a As ArrayList
'>
'>    Set a = lst
'>
'>    a.CopyToListObject ActiveSheet.ListObjects(2)
'>
'>End Sub
'>
'>```
'>
'>**See also**
'>
'>* Dictionary.CopyToListObject
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListObject(LO As ListObject, KeyField As String, Optional obj As Object = Nothing) As IDictionary

    Dim lst As IDictionary
    Dim objRec As Object
    Dim i As Long
    Dim j As Long
    Dim Key As Variant
        
    
    '1次元目はArrayList固定
    Set lst = New Dictionary
        
    For i = 1 To LO.DataBodyRange.Rows.Count
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = Core.GetNewInstance(New Dictionary)
            Case Else
                Set objRec = Core.GetNewInstance(obj)
        End Select
        
        For j = 1 To LO.HeaderRowRange.Columns.Count
            Core.CastICompatibleProperty(objRec).SetProperty LO.HeaderRowRange(, j).Value, LO.DataBodyRange.Cells(i, j).Value
        Next
    
        'Keyフィールド取得
        j = LO.ListColumns(KeyField).index
        Key = LO.DataBodyRange.Cells(i, j).Value
        
        If lst.ContainsKey(Key) Then
            Err.Raise 512 + 1, TypeName(Me) & ".ParseFromListObject", "Duplicate Key"
        End If
        
        lst.Add Key, objRec
    
    Next

    Set ParseFromListObject = lst

End Function
'>---
'>#### CopyToListObject(Dictionary)
'>
'>**Syntax**
'>
'>```
'>Dictionary.CopyToListObject objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Object|copyする ListBox or ComboBox を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ListObject(テーブル機能)へコピー<br>
'>ArrayListの内容をListobjectにコピーする。<br>
'>ListobjectにDictionaryのレコードを対応づける列名が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>Sub Dictionary_ParseFromListObject_Sample()
'>
'>    Dim lst As IDictionary
'>    Dim dic As IDictionary
'>    Dim Key As Variant
'>
'>    Set lst = Dictionary.ParseFromListObject(ActiveSheet.ListObjects(1))
'>
'>    For Each dic In lst
'>
'>        For Each Key In dic.Keys
'>
'>            Debug.Print dic.Item(Key)
'>
'>        Next
'>
'>    Next
'>
'>    Dim a As ArrayList
'>
'>    Set a = lst
'>
'>    a.CopyToListObject ActiveSheet.ListObjects(2)
'>
'>End Sub
'>```
'>
'>**See also**
'>
'>* Dictionary.ParseFromListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToListObject(LO As ListObject)

    Dim lst As IDictionary
    Dim Key As Variant
    Dim obj As Object
    Dim v As Variant

    '1次元目はDictionary固定
    Set lst = Me

    'リストオブジェクトのクリア
    If Not LO.DataBodyRange Is Nothing Then
        LO.DataBodyRange.ClearContents
        LO.DataBodyRange.Delete
    End If
    
    Dim i As Long
    Dim j As Long
    
    i = 1
    For Each Key In lst.Keys
        
        Set obj = lst.Item(Key)
        
        LO.ListRows.Add
        For j = 1 To LO.HeaderRowRange.Columns.Count

            LO.DataBodyRange.Cells(i, j).Value = Core.CastICompatibleProperty(obj).GetProperty(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
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
    If VBA.IsObject(mDic.Item(KeyValue)) Then
        Set IDictionary_Item = mDic.Item(KeyValue)
    Else
        IDictionary_Item = mDic.Item(KeyValue)
    End If
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
    mDic.Key(OldKey) = NewKey
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
    IDictionary_ToString = ToString()
End Function
Private Function ToString() As String

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
                
                SB.Append """" & KeyValue & """:" & v.ToString
            
            Case IsEmpty(v)
                
                SB.Append """" & KeyValue & """:" & "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        SB.Append """" & KeyValue & """:" & LCase(CStr(v))
        
                    Case vbString
                    
                        SB.Append """" & KeyValue & """:" & Escape(v)
                    
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
' 制御文字変換
Private Function Escape(ByVal v As String) As String

    Escape = """" & Convert.Escape(v) & """"

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
    
    Dim d As IDictionary
    Set d = New Dictionary
    
    Set INewInstance_NewInstance = d

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
Private Sub ICompatibleProperty_SetProperty(prop As Variant, Val As Variant)
    
    Dim dic As IDictionary
    
    Set dic = Me.GetIDictionary
    
    If dic.ContainsKey(prop) Then
        If IsObject(Val) Then
            Set dic.Item(prop) = Val
        Else
            dic.Item(prop) = Val
        End If
    Else
        dic.Add prop, Val
    End If
    
End Sub
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
Private Function ICompatibleProperty_GetProperty(prop As Variant) As Variant
    If IsObject(Me.GetIDictionary.Item(prop)) Then
        Set ICompatibleProperty_GetProperty = Me.GetIDictionary.Item(prop)
    Else
        Let ICompatibleProperty_GetProperty = Me.GetIDictionary.Item(prop)
    End If
End Function
