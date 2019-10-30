VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoreEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2
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
'>### CoreEx クラス
'>
'>**Remarks**
'>
'>CoreExクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>ArrayList/Dictionary等のオブジェクトに依存する便利メソッド等<br>
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### CastICompatibleProperty
'>
'>**Syntax**
'>
'>```
'>Set obj = Convert.CastICompatibleProperty(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|ICompatiblePropertyに対応したオブジェクト|
'>
'>**Return Value**
'>
'>ICompatiblePropertyにキャストされたオブジェクト
'>
'>**Remarks**
'>
'>ICompatibleProperty変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICompatibleProperty
'>
Public Function CastICompatibleProperty(ByRef obj As Object) As ICompatibleProperty
    Set CastICompatibleProperty = obj
End Function
'-------------------------------------------------
' NewInstance
'-------------------------------------------------
Public Function GetNewInstance(obj As INewInstance) As Object
    Set GetNewInstance = obj.NewInstance
End Function
'>---
'>#### IlistToListbox(CoreEx)
'>
'>**Syntax**
'>
'>```
'>CoreEx.IListToListbox list, objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|IList|copy元のIList I/F を指定|
'>|objList|必須|Object|copyする ListBox or ComboBox を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ListBox/ComboBox へコピー<br>
'>ArrayList/LinkedListの内容をListBox/ComboBoxにコピーする。<br>
'>ListBox/ComboBoxのTagにArrayListのレコードを対応づけるフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    'これはコードではなくプロパティウィンドウで設定しても良い
'>    lstFavorite.Tag = "No,Name,Address"
'>
'>    CoreEx.IListToListbox list, lstFavorite
'>
'>```
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub IListToListbox(ary As IList, objList As Object)

    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant

    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".IListToListbox", "Item name is not set in tag"
    End If

    '配列に変換
    Fields = Split(objList.Tag, ",")

    'リストボックスのクリア
    objList.Clear

    Dim i As Long
    Dim j As Long
    i = 0
    For Each objRec In ary
        j = 0
        objList.AddItem
        For Each v In Fields

            '設定されているフィールドのみコピーする。
            If v <> "" Then
                objList.list(i, j) = CoreEx.CastICompatibleProperty(objRec).GetProperty(v)
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### IListToListObject(CoreEx)
'>
'>**Syntax**
'>
'>```
'>CoreEx.IListToListObject list, objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|Ilist|copy元のIList I/F を指定|
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
'>ListobjectにArrayListのレコードを対応づける列名が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub IListToListObject(lst As IList, LO As ListObject)

    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant

    'リストオブジェクトのクリア
    If Not LO.DataBodyRange Is Nothing Then
        LO.DataBodyRange.ClearContents
        LO.DataBodyRange.Delete
    End If
    
    Dim i As Long
    Dim j As Long
    
    i = 1
    For Each objRec In lst
        
        LO.ListRows.Add
        For j = 1 To LO.HeaderRowRange.Columns.Count

            LO.DataBodyRange.Cells(i, j).Value = CoreEx.CastICompatibleProperty(objRec).GetProperty(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### IListToRange(CoreEx)
'>
'>**Syntax**
'>
'>```
'>CoreEx.IListToRange list, objct
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|IList|copy元のIList I/F を指定|
'>|object|必須|Range|copyするシートの範囲(Range)の左上を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ArrayListをシートへ貼り付ける<br>
'>指定されたRangeを起点にArrayListの内容を貼り付ける<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub IListToRange(lst As IList, TopLeft As Range)

    Dim v As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    v = lst.ToArray
    
    If Not Core.TryUBound(v, lngRow, 1) Then
        Exit Sub
    End If
    If Not Core.TryUBound(v, lngCol, 2) Then
        Exit Sub
    End If

    TopLeft.resize(lngRow + 1, lngCol + 1).Value = v

End Sub
'>---
'>#### IDictionaryToListbox(CoreEx)
'>
'>**Syntax**
'>
'>```
'>CoreEx.IDictionaryToListbox dic, objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|dic|必須|Object|copy元のIDictionary I/Fを指定|
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
'>    CoreEx.IDictionaryToListbox dic, lstFavorite
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
Public Sub IDictionaryToListbox(ary As IDictionary, objList As Object)

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
                objList.list(i, j) = CoreEx.CastICompatibleProperty(objRec).GetProperty(v)
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### IDictionaryToListObject(CoreEx)
'>
'>**Syntax**
'>
'>```
'>CoreEx.IDictionaryToListObject dic, objList
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|dic|必須|Object|copy元のIDictionary I/Fを指定|
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
'>**See also**
'>
'>* Dictionary
'>* OrderdDictionary
'>* SortedDictionary
'>* MRecord
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub IDictionaryToListObject(lst As IDictionary, LO As ListObject)

    Dim Key As Variant
    Dim obj As Object
    Dim v As Variant

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

            LO.DataBodyRange.Cells(i, j).Value = CoreEx.CastICompatibleProperty(obj).GetProperty(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
