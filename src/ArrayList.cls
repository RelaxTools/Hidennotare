VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
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
'>### ArrayList クラス
'>
'>**Remarks**
'>
'>* このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'>* Collection をベースに実装しているが、添字は0からなので注意
'>* 通常使用するメソッドはほとんど IList インターフェイス側にあるので IList の変数に代入して使うことになります。
'>
'>**Collection と ArrayList との比較**
'>
'>|No.|Property/Method|Collection|ArrayList(本クラス)|
'>|---|---|---|---|
'>|1|Add|〇|AddとInsertに分割|
'>|2|Item|〇|〇|
'>|3|Remove|〇|RemoveAt|
'>|4|Count|〇|〇|
'>|5|Sort|×|〇|
'>|6|ToArray|×|〇|
'>|7|ToCollection|×|〇|
'>|8|ToString|×|〇|
'>
'>**Implements**
'>* IList
'>* ICursor
'>* IConstructor
'>* INewInstance
'>
'>**Sample**
'>```
'> intUbound = ArrayList.TryUBound(aray)
'>
'> Dim list As IList
'> Set list = New ArrayList
'>
'> Dim list As IList
'> Set list = Constructor(New ArrayList, コレクション)
'>
'> Dim list As IList
'> Set list = Constructor(New ArrayList, １次元配列)
'>```
'>**Dependent**
'>
'>* IConstructor
'>* ICursor
'>* IList
'>* INewInstance
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
Implements INewInstance

Private mCol As Collection
Private mIndex As Long
'--------------------------------------------------------------
' Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mCol = New Collection
    mIndex = 0
End Sub
'--------------------------------------------------------------
' Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'>---
'>#### コンストラクタ(1) Collection
'>
'>**Syntax**
'>
'>```
'>Set ret = Constructor(New ArrayList, コレクション)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|コレクション|必須|Collection|生成したコレクションをArrayListの初期値として設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(ArrayList)|
'>
'>**Remarks**
'>
'>コレクションから ArrayList を生成する場合
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
'>#### コンストラクタ(2) Array
'>
'>**Syntax**
'>
'>```
'>Set ret = Constructor(New ArrayList, Array)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Array|必須|配列|配列を元にArrayListの初期値として設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(ArrayList)|
'>
'>**Remarks**
'>
'>配列から ArrayList を生成する場合
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Dim v As Variant
    Dim i As Long
    
    Select Case Args.Count
        
        '引数なし
        Case 0
        
        '引数１個
        Case 1
            Select Case True
                
                'コレクションからの変換
                Case TypeOf Args(1) Is Collection
                    Call IList_Clear
                    Set mCol = Args(1)
                
                '１次元配列からの変換
                Case IsArray(Args(1))
                    Call IList_Clear
                    For i = LBound(Args(1)) To UBound(Args(1))
                        mCol.Add Args(1)(i)
                    Next
                Case Else
                    'エラー
                    Exit Function
            End Select
        
        '引数からの追加
        Case Else
            Call IList_Clear
            For Each v In Args
                mCol.Add v
            Next
    End Select
    
    Set IConstructor_Instancing = Me
    
End Function
'>---
'>#### IsArray(ArrayList)
'>
'>**Syntax**
'>
'>```
'>flg = ArrayList.IsArray(v)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|v|必須|Variant|配列かどうかチェックしたい変数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|配列である:True / 配列ではないまたはエラー:False|
'>
'>**Remarks**
'>
'>空配列対応 IsArray<br>
'>配列の場合 True を配列じゃない場合またはエラーが発生した場合は Falseを戻します。<br>
'>そもそもの話ですが、Dim v As Variant で v = Array() で初期値を指定すれば通常のIsArrayで事足ります。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function IsArray(v As Variant) As Boolean

    On Error GoTo e

    IsArray = False

    If VBA.IsArray(v) Then
        IsArray = IIf(UBound(v) < LBound(v), False, True)
    Else
        IsArray = False
    End If

e:
    Exit Function

End Function
'>---
'>#### TryUbound(ArrayList)
'>
'>**Syntax**
'>
'>```
'>flg = ArrayList.TryUBound(ary, idx, [dimension = 1])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|ary|必須|Variant|チェックをする配列を指定|
'>|idx|必須|Long|配列の上限を返却する変数を指定|
'>|dimension|省略可|Long|省略時１（次元）。テストする次元数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|正常:True / エラー:False|
'>
'>**Remarks**
'>
'>エラー検知付きUBound<br>
'>テストする配列が１次元の場合に２次元の上限を求めようとすると実行時エラーになる。<br>
'>そのためエラー検知可能なUboundとして作成。エラーの場合には False が返却される。
'>
'>**Example**
'>
'>**See also**
'>
'>* ArrayList.TryLBound
'>
Public Function TryUBound(ary As Variant, idx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    idx = UBound(ary, dimension)

    TryUBound = True
    
    Exit Function
e:
    TryUBound = False

End Function
'>---
'>#### TryLBound(ArrayList)
'>
'>**Syntax**
'>
'>```
'>flg = ArrayList.TryLBound(ary, idx, [dimension = 1])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|ary|必須|Variant|チェックをする配列を指定|
'>|idx|必須|Long|配列の下限を返却する変数を指定|
'>|dimension|省略可|Long|省略時１（次元）。テストする次元数を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|正常:True / エラー:False|
'>
'>**Remarks**
'>
'>エラー検知付きUBound<br>
'>テストする配列が１次元の場合に２次元の下限を求めようとすると実行時エラーになる。<br>
'>そのためエラー検知可能なUboundとして作成。エラーの場合には False が返却される。
'>
'>**Example**
'>
'>**See also**
'>
'>* ArrayList.TryUBound
'>
Public Function TryLBound(ary As Variant, lngIdx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    lngIdx = LBound(ary, dimension)

    TryLBound = True
    
    Exit Function
e:
    TryLBound = False

End Function
'>---
'>#### ParseFromListbox(ArrayList)
'>
'>**Syntax**
'>
'>```
'>Set ret = ArrayList.ParseFromListbox(objList,[obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Object|Parseする ListBox or ComboBox を指定|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|新しいインスタンス(ArrayList)|
'>
'>
'>**Remarks**
'>
'>ListBox/ComboBox パース<br>
'>ListBox/ComboBoxからArrayList(レコードはDictionary)を生成する。<br>
'>Tagにレコードを生成するためのフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    lstFavorite.Tag = "No,Name,Address" 'これはコードではなくプロパティウィンドウで設定しても良い
'>
'>    Dim lst As IList
'>    Set lst = ArrayList.ParseFromListbox(lstFavorite, New FavoriteBean)
'>
'>```
'>
'>**See also**
'>
'>* ArrayList.CopyToListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListbox(objList As Object, Optional obj As Object = Nothing) As IList

    Dim ary As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim i As Long
    Dim j As Long
    
    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".ParseFromListbox", "Item name is not set in tag"
    End If
    
    '配列に変換
    Fields = Split(objList.Tag, ",")
    
    '1次元目はArrayList固定
    Set ary = New ArrayList
        
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
            End If
        Next
            
        ary.Add objRec
    
    Next

    Set ParseFromListbox = ary

End Function
'>---
'>#### CopyToListbox(ArrayList)
'>
'>**Syntax**
'>
'>```
'>ArrayList.CopyToListbox objList
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
'>ListBox/ComboBoxのTagにArrayListのレコードを対応づけるフィールド名リスト（カンマ区切り）が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>    lstFavorite.Tag = "No,Name,Address" 'これはコードではなくプロパティウィンドウで設定しても良い
'>
'>    ArrayList.CopyToListbox lstFavorite
'>
'>```
'>
'>**See also**
'>
'>* ArrayList.ParseFromListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToListbox(objList As Object)

    Dim ary As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant

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
    For Each objRec In ary
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
'>#### ParseFromListObject(ArrayList)
'>
'>**Syntax**
'>
'>```
'>Set ret = ArrayList.ParseFromListObject(objList,[obj = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|objList|必須|Listobject|Parseする Listobject を指定|
'>|obj|省略可|Object|レコードに使用するオブジェクトを指定。省略時は Dictionary |
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Boolean|新しいインスタンス(ArrayList)|
'>
'>**Remarks**
'>
'>ListObject パース<br>
'>ListobjectからArrayList(レコードはDictionary)を生成する。<br>
'>ListObject(テーブル)の列ヘッダの情報が必須<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>Sub ArrayList_ParseFromListObject_Sample()
'>
'>    Dim lst As IList
'>    Dim dic As IDictionary
'>    Dim Key As Variant
'>
'>    Set lst = ArrayList.ParseFromListObject(ActiveSheet.ListObjects(1))
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
'>* ArrayList.CopyToListObject
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Function ParseFromListObject(LO As ListObject, Optional obj As Object = Nothing) As IList

    Dim lst As IList
    Dim objRec As Object
    Dim i As Long
    Dim j As Long
    
    '1次元目はArrayList固定
    Set lst = New ArrayList
        
    For i = 1 To LO.DataBodyRange.Rows.Count
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = Core.GetNewInstance(New Dictionary)
            Case Else
                Set objRec = Core.GetNewInstance(obj)
        End Select
        
        For j = 1 To LO.HeaderRowRange.columns.Count
        
            Core.CastICompatibleProperty(objRec).SetProperty LO.HeaderRowRange(, j).Value, LO.DataBodyRange.Cells(i, j).Value
        Next
    
        lst.Add objRec
    
    Next

    Set ParseFromListObject = lst

End Function
'>---
'>#### CopyToListObject(ArrayList)
'>
'>**Syntax**
'>
'>```
'>ArrayList.CopyToListObject objList
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
'>ListobjectにArrayListのレコードを対応づける列名が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>Sub ArrayList_ParseFromListObject_Sample()
'>
'>    Dim lst As IList
'>    Dim dic As IDictionary
'>    Dim Key As Variant
'>
'>    Set lst = ArrayList.ParseFromListObject(ActiveSheet.ListObjects(1))
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
'>* ArrayList.ParseFromListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToListObject(LO As ListObject)

    Dim lst As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant

    '1次元目はArrayList固定
    Set lst = Me

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
        For j = 1 To LO.HeaderRowRange.columns.Count

            LO.DataBodyRange.Cells(i, j).Value = Core.CastICompatibleProperty(objRec).GetProperty(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### CopyToRange(ArrayList)
'>
'>**Syntax**
'>
'>```
'>ArrayList.CopyToRange objct
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
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
'>ListobjectにArrayListのレコードを対応づける列名が必須。<br>
'>レコードに使用するオブジェクトは INewInstance と ICompatibleProperty を Implements したクラスであること。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ArrayList.ParseFromListbox
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToRange(TopLeft As Range)

    Dim v As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    v = IList_ToArray
    
    If Not TryUBound(v, lngRow, 1) Then
        Exit Sub
    End If
    If Not TryUBound(v, lngCol, 2) Then
        Exit Sub
    End If

    TopLeft.resize(lngRow + 1, lngCol + 1).Value = v

End Sub
'>---
'>#### GetIList(ArrayList)
'>
'>**Syntax**
'>
'>```
'> Set obj = ArrayList.GetIList()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|object|必須|Object|IList I/F に対応するクラス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|obj|IList|現在のインスタンス(IList I/F)|
'>
'>**Remarks**
'>
'>IList I/F を取得する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Public Function GetIList() As IList
    Set GetIList = Me
End Function
'>---
'>#### Clear(IList I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Clear
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList I/F|IList I/Fのインスタンス|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayListの全項目をクリアします。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Sub IList_Clear()

    Set mCol = New Collection
    mIndex = 0
    
End Sub
'>---
'>#### Add(IList I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Add(var)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>|var|必須|Variant|追加する変数を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList に項目を追加する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Sub IList_Add(obj As Variant)
    
    mCol.Add obj

End Sub
'>---
'>#### Insert(IList I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Insert(index, var)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>|index|必須|long|追加する位置を指定します。|
'>|var|必須|Variant|追加する変数を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList の指定位置に項目を追加する。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Sub IList_Insert(idx As Long, v As Variant)

    If idx < 0 Or idx >= mCol.Count Then
        Err.Raise 512 + 1, TypeName(Me) & ".Insert", "IndexOutOfRangeException"
    End If
    
    mCol.Add v, , idx + 1

End Sub
'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = mCol.[_NewEnum]
End Function
'>---
'>#### Count(IList I/F)
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
'>|expression|必須|IList|IList I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Long|ArrayListの要素数を返却|
'>
'>**Remarks**
'>
'>IList I/F のプロパティです。<br>
'>ArrayList 要素数を返却
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Property Get IList_Count() As Long
    IList_Count = mCol.Count
End Property
'>---
'>#### Item(IList I/F)
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
'>|expression|必須|IList|IList I/F のインスタンス|
'>|idx|必須|IList|Long|取得する項目の位置を指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|ArrayList の指定位置の項目|
'>
'>**Remarks**
'>
'>IList I/F のプロパティです。<br>
'>ArrayList の指定位置の項目を取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Property Get IList_Item(idx As Long) As Variant
    
    If idx < 0 Or idx >= mCol.Count Then
        Err.Raise 512 + 1, TypeName(Me) & ".Item", "IndexOutOfRangeException"
    End If

    If VBA.IsObject(mCol.Item(idx + 1)) Then
        Set IList_Item = mCol.Item(idx + 1)
    Else
        IList_Item = mCol.Item(idx + 1)
    End If
    
End Property
'>---
'>#### RemoveAt(IList I/F)
'>
'>**Syntax**
'>
'>```
'>expression.RemoveAt(idx)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>|idx|必須|IList|Long|削除する項目の位置を指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList の指定位置の項目を削除
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Sub IList_RemoveAt(idx As Long)
    
    If idx < 0 Or idx >= mCol.Count Then
        Err.Raise 512 + 1, TypeName(Me) & ".RemoveAt", "IndexOutOfRangeException"
    End If
    
    mCol.Remove idx + 1
End Sub
'>---
'>#### Sort(IList I/F)
'>
'>**Syntax**
'>
'>```
'>expression.Sort([cp = Nothing])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>|CP|省略可|IComparer|ソート条件を指定します。省略時はBinaryCompareです。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList の指定位置の項目を削除
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Sub IList_Sort(Optional ByVal CP As IComparer = Nothing)

    Dim i As Long
    Dim j As Long
    Dim col2 As Collection
    Dim blnFind As Boolean
    
    If CP Is Nothing Then
        'Interfaceも普通のクラスなのでDefault比較として利用
        Set CP = New IComparer
    End If
    
    'Collectionが空ならなにもしない
    If mCol Is Nothing Then
        Exit Sub
    End If

    'Collectionの要素数が０または１の場合ソート不要
    If mCol.Count <= 1 Then
        Exit Sub
    End If
    
    Set col2 = New Collection
    
    For i = 1 To mCol.Count
        If col2.Count = 0 Then
            col2.Add mCol(i)
        Else
            blnFind = False
            For j = col2.Count To 1 Step -1
    
                'ファイルの方が大きかった場合、その後に挿入。
                If CP.Compare(mCol(i), col2(j)) >= 0 Then
                    col2.Add mCol(i), , , j
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                col2.Add mCol(i), , 1
            End If
        End If
    
    Next
    
    Set mCol = col2
    Set col2 = Nothing

End Sub
'>---
'>#### ToArray(IList I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToArray()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Variant|配列を返却|
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList を１～２次元配列に変換
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Function IList_ToArray() As Variant

    Dim v As Variant
    Dim s As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    Dim j As Long
    Dim result As Variant
    Dim lst As IList
    
    result = Array()
    
    lngRow = mCol.Count
    lngCol = 0
    
    Set lst = Me
    
    'Dictionary or ArrayList の２次元を想定
    For Each v In lst
        If IsDictionary(v) Or IsList(v) Then
            If lngCol < v.Count Then
                lngCol = v.Count
            End If
        End If
    Next

    Select Case True
        Case lngRow = 0
    
            IList_ToArray = Empty

        '1Dの場合
        Case lngRow > 0 And lngCol = 0
        
            ReDim result(0 To lngRow - 1)
        
            For i = 0 To lst.Count - 1
                If IsObject(lst(i)) Then
                    Set result(i) = lst(i)
                Else
                    Let result(i) = lst(i)
                End If
            Next
    
            IList_ToArray = result
        
        '2Dの場合
        Case lngRow > 0 And lngCol > 0
        
            ReDim result(0 To lngRow - 1, 0 To lngCol - 1)
        
            For i = 0 To lst.Count - 1
            
                Select Case True
                    'List系の場合
                    Case IsList(lst(i))
                        For j = 0 To lst(i).Count - 1
                            If IsObject(lst(i)(j)) Then
                                Set result(i, j) = lst(i)(j)
                            Else
                                Let result(i, j) = lst(i)(j)
                            End If
                        Next
                    
                    'Dictionary系の場合
                    Case IsDictionary(lst(i))
                        Dim dic As Object
                        Dim Key As Variant
                        
                        Set dic = lst(i)
                        j = 0
                        For Each Key In dic.Keys
                    
                            If IsObject(dic.Item(Key)) Then
                                Set result(i, j) = dic.Item(Key)
                            Else
                                Let result(i, j) = dic.Item(Key)
                            End If
                            
                            j = j + 1
                        Next
                
                End Select
                
            Next
    
            IList_ToArray = result
    End Select
    
End Function
'>---
'>#### ToCollection(IList I/F)
'>
'>**Syntax**
'>
'>```
'>var = expression.ToCollection()
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|IList|IList I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|Collection|Collectionを返却|
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList をCollectionに変換
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Function IList_ToCollection() As Collection
    
    Set IList_ToCollection = mCol

End Function
'>---
'>#### ToString(IList I/F)
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
'>|expression|必須|IList|IList I/F のインスタンス|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|String|JSON文字列を返却|
'>
'>**Remarks**
'>
'>IList I/F のメソッドです。<br>
'>ArrayList をJSONに変換
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Private Function IList_ToString() As String

    Dim v As Variant
    Dim SB As StringBuilder
    Set SB = New StringBuilder
    
    For Each v In mCol
    
        Select Case True
            Case IsObject(v)
                
                SB.Append v.ToString
            
            Case IsEmpty(v)
                
                SB.Append "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        SB.Append LCase(CStr(v))
        
                    Case vbString
                    
                        SB.Append Escape(v)
                    
#If Win64 Then
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbLongLong
#Else
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
#End If
                        SB.Append CStr(v)
        
                    Case Else
        
                        'エラー
                        Err.Raise 512 + 1, TypeName(Me) & ".ToString", "don't cast type (" & TypeName(v) & ")"
        
                End Select
        End Select
        
    Next

    IList_ToString = SB.ToString(", ", "[", "]")

End Function
'--------------------------------------------------------------
' 制御文字変換
'---------------------------------------------------------------------
Private Function Escape(ByVal v As String) As String

    Escape = """" & Convert.Escape(v) & """"

End Function
'>---
'>#### GetCursor(IList I/F)
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
'>|expression|必須|IList|IList I/F を指定します。|
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
'>* IList
'>* ICursor
'>
Private Property Get IList_GetCursor() As ICursor
    Set IList_GetCursor = Me
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
    ICursor_Eof = mIndex >= mCol.Count
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
    
    If mIndex >= mCol.Count Then
        If IsObject(mCol.Item(mCol.Count)) Then
            Set ICursor_Item = mCol(mCol.Count)
        Else
            ICursor_Item = mCol.Item(mCol.Count)
        End If
    Else
        If IsObject(mCol.Item(mIndex + 1)) Then
            Set ICursor_Item = mCol.Item(mIndex + 1)
        Else
            ICursor_Item = mCol.Item(mIndex + 1)
        End If
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
    If mIndex > mCol.Count Then
        mIndex = mCol.Count
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
'>|var|Object|新しいArrayListのインスタンス(IList I/F)|
'>
'>**Remarks**
'>
'>INewInstance I/F のメソッドです。<br>
'>新しいインスタンス(IList I/F)を生成して返却します。
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>* INewInstance
'>
Private Function INewInstance_NewInstance() As Object

    Dim lst As IList
    
    Set lst = New ArrayList
    
    Set INewInstance_NewInstance = lst

End Function

