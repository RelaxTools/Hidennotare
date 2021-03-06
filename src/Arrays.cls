VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
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
'>### Arrays クラス
'>
'>**Remarks**
'>
'>Arraysクラス<br>
'>配列に関係する操作を行うクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
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
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (destination As Any, ByVal Length As LongPtr)

'-----------------------------------------------------------------------------------------------------
'>---
'>#### ToCollection
'>
'>**Syntax**
'>
'>```
'>Set col = Arrays.ToCollection(args)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|args|必須|Variant|チェックをする配列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|col|Collection|Collectionを返却する|
'>
'>**Remarks**
'>
'>配列またはParamArrayをCollectionに変換<br>
'>
'>**Example**
'>
'-----------------------------------------------------------------------------------------------------
Public Function ToCollection(Args As Variant) As Collection
        
    Dim col As Collection
    Dim v As Variant
    
    Set col = New Collection
        
    For Each v In Args
        col.Add v
    Next

    Set ToCollection = col

End Function
'-----------------------------------------------------------------------------------------------------
'>---
'>#### ToList
'>
'>**Syntax**
'>
'>```
'>Set col = Arrays.ToList(args)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|args|必須|Variant|変換する配列またはCollectionを指定|
'>|obj|任意|INewInstance|変換後のIList I/F 省略した場合、ArrayList|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|col|Collection|Collectionを返却する|
'>
'>**Remarks**
'>
'>配列またはCollectionをArrayListに変換<br>
'>
'>**Example**
'>
'-----------------------------------------------------------------------------------------------------
Public Function ToList(ary As Variant, Optional ByVal obj As INewInstance = Nothing) As IList

    Dim lst As IList
    Dim v As Variant
       
    'ArrayList 固定
    If obj Is Nothing Then
        Set lst = ArrayList.NewInstance()
    Else
        Set lst = obj.NewInstance()
    End If
    
    If TypeOf ary Is Collection Or IsArray(ary) Then
    
        For Each v In ary
            lst.Add v
        Next
    
    End If
    
    Set ToList = lst

End Function
'>---
'>#### TryUbound
'>
'>**Syntax**
'>
'>```
'>flg = Arrays.TryUBound(ary, idx, [dimension = 1])
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
'>* Arrays.TryLBound
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
'>#### TryLBound
'>
'>**Syntax**
'>
'>```
'>flg = Arrays.TryLBound(ary, idx, [dimension = 1])
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
'>* Arrays.TryUBound
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
'>#### Length
'>
'>**Syntax**
'>
'>```
'>length = Arrays.Length(ary)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|ary|必須|Variant|チェックをする配列を指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|length|Long|配列サイズを返却|
'>
'>**Remarks**
'>
'>配列のサイズを調べる<br>
'>指定された配列のサイズを返却する<br>
'>
'>**Example**
'>
'>**See also**
'>
'>* Arrays.TryLBound
'>* Arrays.TryUBound
'>
Function Length(var As Variant) As Long

    Dim lngValue As Long
    Dim lngMin As Long
    Dim lngMax As Long
    
    If IsArray(var) Then
    
        If Me.TryLBound(var, lngValue) Then
            lngMin = lngValue
        End If
        
        If Me.TryUBound(var, lngValue) Then
            lngMax = lngValue
        End If
    
        Length = lngMax - lngMin + 1
    
    Else
        Err.Raise 512 + 1, TypeName(Me) & ".", "support array only"
    End If

End Function
'>---
'>#### CopyToListbox(Arrays)
'>
'>**Syntax**
'>
'>```
'>Arrays.CopyToListbox list, obj, Fields
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|IList|copy元のIList I/F を指定|
'>|obj|必須|Object|copyする ListBox or ComboBox を指定|
'>|Fields|必須|フィールド名リスト（配列またはコレクションまたはIList）|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ListBox/ComboBox へコピー<br>
'>ArrayList/LinkedListの内容をListBox/ComboBoxにコピーする。<br>
'>ListBox/ComboBoxのTagにArrayListのレコードを対応づけるフィールド名リストが必須。<br>
'>レコードに使用するオブジェクトは INewInstance と IValueObject を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>
'>    Arrays.CopyToListbox list, lstFavorite, Array(No, Name, Address)
'>
'>```
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* IValueObject
'>
Public Sub CopyToListbox(ary As IList, objList As Object, Fields As Variant)

    Dim objRec As Object
    Dim v As Variant

    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".CopyToListbox", "Item name is not set in tag"
    End If

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
                objList.list(i, j) = Objects.AsIValueObject(objRec).Fields(v)
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### CopyToListObject(Arrays)
'>
'>**Syntax**
'>
'>```
'>Arrays.CopyToListObject list, objList
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
'>レコードに使用するオブジェクトは INewInstance と IValueObject を Implements したクラスであること。
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* ICompatibleProperty
'>
Public Sub CopyToListObject(lst As IList, LO As ListObject)

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

            LO.DataBodyRange.Cells(i, j).Value = Objects.AsIValueObject(objRec).Fields(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### CopyToRange(Arrays)
'>
'>**Syntax**
'>
'>```
'>Arrays.CopyToRange list, range
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|IList|copy元のIList I/F を指定|
'>|range|必須|Range|copyするシートの範囲(Range)の左上を指定|
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
Public Sub CopyToRange(lst As IList, TopLeft As Range)

    Dim v As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    v = lst.ToArray
    
    If Not Arrays.TryUBound(v, lngRow, 1) Then
        Exit Sub
    End If
    If Not Arrays.TryUBound(v, lngCol, 2) Then
        Exit Sub
    End If

    TopLeft.resize(lngRow + 1, lngCol + 1).Value = v

End Sub
'>---
'>#### CopyToListView(Arrays)
'>
'>**Syntax**
'>
'>```
'>Arrays.CopyToListView list, obj, Fields
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|list|必須|IList|copy元のIList I/F を指定|
'>|obj|必須|Object|copyする ListView を指定|
'>|Fields|必須|フィールド名リスト（配列またはコレクションまたはIList）|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ListView へコピー<br>
'>ArrayList/LinkedListの内容をListViewにコピーする。<br>
'>ArrayListの中のIDictionay/IValueObjectに対応づけるフィールド名リストが必須。<br>
'>レコードに使用するオブジェクトは INewInstance と IValueObject を Implements したクラスであること。
'>
'>**Example**
'>
'>```
'>
'>    Arrays.CopyToListView list, lvwFavorite, Array(No, Name, Address)
'>
'>```
'>
'>**See also**
'>
'>* ArrayList
'>* LinkedList
'>* Dictionary
'>* INewInstance
'>* IValueObject
'>
Public Sub CopyToListView(ary As IList, objList As Object)

    Dim objRec As Object
    Dim v As Variant

    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Err.Raise 512 + 1, TypeName(Me) & ".IListToListbox", "Item name is not set in tag"
    End If
    
    Dim Fields As Variant
    
    Fields = Split(objList.Tag, ",")

    'ListViewのクリア
    objList.ListItems.Clear

    Dim i As Long
    Dim j As Long
    i = 0
    For Each objRec In ary
        j = 0
        For Each v In Fields

            '設定されているフィールドのみコピーする。
            If v <> "" Then
                If j = 0 Then
                    With objList.ListItems.Add
                        .Text = Objects.AsIValueObject(objRec).Fields(v)
                    End With
                Else
                    objList.ListItems(i + 1).SubItems(j) = Objects.AsIValueObject(objRec).Fields(v)
                
                End If
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub
'>---
'>#### ArrayMove(Arrays)
'>
'>**Syntax**
'>
'>```
'>Arrays.ArrayMove src, srcPos, dstPos, length
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|src|必須|処理を行う配列を指定|
'>|srcPos|必須|コピー元のIndexを指定|
'>|dstPos|必須|コピー先のIndexを指定|
'>|length|必須|コピーする配列の数|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>配列内のデータ移動<br>
'>ArrayListのInsertやRemoveAtで配列を詰めなおすときなどに用いる<br>
'>配列が重なっていても動作。移動して空いた場所にはEmptyが設定される。<br>
'>
'>**See also**
'>
'>* ArrayList
'>
Sub ArrayMove(src As Variant, srcPos As Long, dstPos As Long, Length As Long)

#If Win64 Then
    Const VARIANT_SIZE As Long = 24
#Else
    Const VARIANT_SIZE As Long = 16
#End If

    Dim zeroPos As Long
    Dim zeroLen As Long
    
    If srcPos > dstPos Then
        zeroPos = dstPos + Length
        zeroLen = srcPos - dstPos
    Else
        zeroPos = srcPos
        zeroLen = dstPos - srcPos
    End If

    CopyMemory src(dstPos), src(srcPos), VARIANT_SIZE * Length
    ZeroMemory src(zeroPos), VARIANT_SIZE * zeroLen

End Sub
Public Function IsArray(arg As Variant) As Boolean

    On Error GoTo e

    IsArray = VBA.IsArray(arg)

    Exit Function
e:

    IsArray = False

End Function

