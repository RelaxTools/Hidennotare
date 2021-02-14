VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedList"
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
'>### LinkedList クラス
'>
'>**Remarks**
'>
'>- Collection をベースに実装しているが、添字は0からなので注意。
'>- 行途中の追加、削除に強いですが、添え字によるランダムアクセスが弱いです。
'>- 通常使用するメソッドはほとんど IList インターフェイス側にあるので IList の変数に代入して使うことになります。
'>- ICursor インターフェースに対応しており他の ICursor I/F と同様にループ可能。
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Collection と LinkedList との比較**
'>
'>|No.|Property/Method|Collection|LinkedList(本クラス)|
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
'>* INewInstance
'>
'>**Sample**
'>```
'> intUbound = LinkedList.TryUBound(aray)
'>
'> Dim list As IList
'> Set list = New LinkedList
'>
'> Dim list As IList
'> Set list = LinkedList.NewInstance(コレクション)
'>
'> Dim list As IList
'> Set list = LinkedList.NewInstance(１次元配列)
'>```
'>**Dependent**
'>
'>* ICursor
'>* IList
'>* INewInstance
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-08-08|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IList
Implements INewInstance

Private mCol As Collection
Private mIndex As Long

'>---
'>#### コンストラクタ(1) Collection
'>
'>**Syntax**
'>
'>```
'>Set ret = LinkedList.NewInstance(collection)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|collection|必須|Collection|生成したコレクションをLinkedListの初期値として設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(LinkedList)|
'>
'>**Remarks**
'>
'>コレクションから LinkedList を生成する場合
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
'>Set ret = LinkedList.NewInstance(Array)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|Array|必須|配列|配列を元にLinkedListの初期値として設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(LinkedList)|
'>
'>**Remarks**
'>
'>配列から LinkedList を生成する場合
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function NewInstance(ParamArray Args() As Variant) As IList

    Dim IL As IList
    Dim var As Variant
    Dim v As Variant
    
    var = Args
    
    Set IL = New LinkedList
    IL.Clear
    
    Select Case Arrays.Length(var)
        Case 0
        Case 1
        
            Select Case True
                
                'コレクションまたは配列からの変換
                Case TypeOf Args(0) Is Collection, IsArray(Args(0))
                    
                    For Each v In Args(0)
                        IL.Add v
                    Next
                
                Case Else
                
                    IL.Add Args(0)
            
            End Select
            
        Case Else
            
            For Each v In Args
                IL.Add v
            Next
    
    End Select

    Set NewInstance = IL

End Function
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
'>LinkedListの全項目をクリアします。
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
'>LinkedList に項目を追加する。
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
'>LinkedList の指定位置に項目を追加する。
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
'>|var|Long|LinkedListの要素数を返却|
'>
'>**Remarks**
'>
'>IList I/F のプロパティです。<br>
'>LinkedList 要素数を返却
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
'>|var|Variant|LinkedList の指定位置の項目|
'>
'>**Remarks**
'>
'>IList I/F のプロパティです。<br>
'>LinkedList の指定位置の項目を取得
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

    Objects.SetWith(IList_Item) = mCol.Item(idx + 1)
    
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
'>LinkedList の指定位置の項目を削除
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
'>LinkedList の指定位置の項目を削除
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
'>LinkedList を１～２次元配列に変換
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
    
    Dim IB As IValueObject
    Dim ID As IDictionary
    Dim IL As IList
    
    
    Set lst = Me
    
    lngRow = lst.Count
    lngCol = 0
    
    result = Array()
    
    'IDictionary or IList の２次元を想定
    For Each v In lst
        
        Select Case True
            Case Objects.InstanceOfIDictionary(v)
                Set ID = v
                If lngCol < ID.Count Then
                    lngCol = ID.Count
                End If
            Case Objects.InstanceOfIList(v)
                Set IL = v
                If lngCol < IL.Count Then
                    lngCol = IL.Count
                End If
            Case Objects.InstanceOfIValueObject(v)
                Set IB = v
                If lngCol < IB.Count Then
                    lngCol = IB.Count
                End If
        End Select
    
    Next

    Select Case True
        Case lngRow = 0
    
            IList_ToArray = Empty

        '1Dの場合
        Case lngRow > 0 And lngCol = 0
        
            ReDim result(0 To lngRow - 1)
        
            For i = 0 To lst.Count - 1
                Objects.SetWith(result(i)) = lst(i)
            Next
    
            IList_ToArray = result
        
        '2Dの場合
        Case lngRow > 0 And lngCol > 0
        
            ReDim result(0 To lngRow - 1, 0 To lngCol - 1)
        
            For i = 0 To lst.Count - 1
            
                Select Case True
                    'List系の場合
                    Case Objects.InstanceOfIList(lst(i))
                    
                        Set IL = lst(i)
                        
                        For j = 0 To IL.Count - 1
                            Objects.SetWith(result(i, j)) = IL(j)
                        Next
                    
                    'Dictionary系の場合
                    Case Objects.InstanceOfIDictionary(lst(i))
                        
                        Dim key As Variant
                        
                        Set ID = lst(i)
                        
                        j = 0
                        For Each key In ID.Keys
                            Objects.SetWith(result(i, j)) = ID.Item(key)
                            j = j + 1
                        Next
                    
                    'ValueObject系の場合
                    Case Objects.InstanceOfIValueObject(lst(i))
                        
                        Set IB = lst(i)
                        
                        j = 0
                        For Each key In IB.FieldNames
                            Objects.SetWith(result(i, j)) = IB.Fields(key)
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
'>LinkedList をCollectionに変換
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
'>LinkedList をJSONに変換
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
    
    With StringBuilder.NewInstance
    
        For Each v In mCol
        
            Select Case True
                Case Objects.InstanceOfIDictionary(v)
                    
                    .Append Objects.AsIDictionary(v).ToString
                
                Case Objects.InstanceOfIList(v)
                    
                    .Append Objects.AsIList(v).ToString
                
                Case Objects.InstanceOfIValueObject(v)
                    
                    .Append Objects.AsIValueObject(v).ToString
                
                Case IsEmpty(v)
                    
                    .Append "null"
                
                Case Else
                    
                    Select Case VarType(v)
                        
                        Case vbBoolean
            
                            .Append LCase(CStr(v))
            
                        Case vbString
                        
                            .Append """" & StringUtils.Escape(v) & """"
                        
#If Win64 Then
                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbLongLong
#Else
                        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
#End If
                            .Append CStr(v)
            
                        Case Else
            
                            'エラー
                            Err.Raise 512 + 1, TypeName(Me) & ".ToString", "don't cast type (" & TypeName(v) & ")"
            
                    End Select
            End Select
            
        Next
    
        IList_ToString = .ToString(", ", "[", "]")

    End With
    
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
        Objects.SetWith(ICursor_Item) = mCol(mCol.Count)
    Else
        Objects.SetWith(ICursor_Item) = mCol.Item(mIndex + 1)
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
'>|var|Object|新しいLinkedListのインスタンス(IList I/F)|
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
    
    Set lst = New LinkedList
    
    Set INewInstance_NewInstance = lst

End Function
'Private Function Escape(ByVal v As String) As String
'
'    Dim strChr As String
'    Dim SB As IStringBuilder
'    Dim i As Long
'
'    Set SB = StringBuilder.NewInstance
'
'    For i = 1 To Len(v)
'
'        strChr = Mid$(v, i, 1)
'
'        Select Case AscW(strChr)
'            Case &H0 To &H7, &HB, &HE To &H1F, &H7F
'                SB.Append "\u" & Hex$(AscW(strChr))
'            Case &H8
'                SB.Append "\b"
'            Case &H9
'                SB.Append "\t"
'            Case &HA
'                SB.Append "\n"
'            Case &HC
'                SB.Append "\f"
'            Case &HD
'                SB.Append "\r"
'            Case &H22
'                SB.Append "\"""
'            Case &H2F
'                SB.Append "\/"
'            Case &H5C
'                SB.Append "\\"
'            Case Else
'                SB.Append strChr
'        End Select
'
'    Next
'
'    Escape = SB.ToString
'
'End Function
