VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
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
'>### ArrayList クラス
'>
'>**Remarks**
'>
'>- Collectionと異なり、添え字は0から始まるので注意。
'>- 行の途中の追加、削除のパフォーマンスが悪いです。添え字によるランダムアクセスに強いです。
'>- 通常使用するメソッドはほとんど IList インターフェイス側にあるので IList の変数に代入して使うことになります。
'>- ICursor インターフェースに対応しており他の ICursor I/F と同様にループ可能。
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
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
'>* INewInstance
'>
'>**Sample**
'>```
'> Dim list As IList
'> Set list = ArrayList.CreateObject()
'>
'> Dim list As IList
'> Set list = ArrayList.CreateObject(コレクション)
'>
'> Dim list As IList
'> Set list = ArrayList.CreateObject(１次元配列)
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
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IList
Implements INewInstance
    
'NewEnum返却用
Private mCol As Collection

Private mArray As Variant
Private mCount As Long
Private mIndex As Long
'>---
'>#### コンストラクタ(1) Collection
'>
'>**Syntax**
'>
'>```
'>Set ret = ArrayList.CreateObject(collection)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|collection|必須|Collection|生成したコレクションをArrayListの初期値として設定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(IList)|
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
'>Set ret = ArrayList.CreateObject(Array)
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
'>|ret|Object|新しいインスタンス(IList)|
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
Public Function CreateObject(ParamArray Args() As Variant) As IList

    Dim cnt As Long
    Dim lngValue As Long
    Dim lngMin As Long
    Dim lngMax As Long
    Dim lngCount As Long
    Dim var As Variant
    Dim v As Variant
    Dim i As Long
    
    var = Args
    
    lngMin = 0
    lngMax = -1
    
    Dim IL As IList
    
    If Core.TryLBound(var, lngValue) Then
        lngMin = lngValue
    End If
    
    If Core.TryUBound(var, lngValue) Then
        lngMax = lngValue
    End If
    
    lngCount = lngMax + 1

    Set IL = New ArrayList
    IL.Clear
    
    Select Case lngCount
        Case 0
        Case 1
        
            Select Case True
                
                'コレクションまたは配列からの変換
                Case TypeOf Args(0) Is Collection, IsArray(Args(0))
                    
                    Core.ArrayToIList Args(0), IL
                
                Case Else
                
                    IL.Add Args(0)
            
            End Select
            
        Case Else
            
            For i = lngMin To lngMax
                IL.Add Args(i)
            Next
    
    End Select

    Set CreateObject = IL
    
End Function
Private Sub Class_Initialize()
    Call IList_Clear
    mIndex = 0
End Sub
Private Sub Class_Terminate()
    Call IList_Clear
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

    ReDim Preserve mArray(0 To mCount)

    Core.SetVar(mArray(mCount)) = obj
    
    mCount = mCount + 1

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
    mArray = Array()
    mCount = 0
End Sub
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
    IList_Count = mCount
End Property
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

    Dim i As Long

    If idx < 0 Or idx >= mCount Then
        Err.Raise 512 + 1, TypeName(Me) & ".Insert", "IndexOutOfRangeException"
    End If
    
    ReDim Preserve mArray(0 To mCount + 1)

    For i = mCount To idx Step -1
        Core.SetVar(mArray(i + 1)) = mArray(i)
    Next
    
    Core.SetVar(mArray(idx)) = v
    
    mCount = mCount + 1

End Sub
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

    If idx < 0 Or idx >= mCount Then
        Err.Raise 512 + 1, TypeName(Me) & ".Item", "IndexOutOfRangeException"
    End If

    Core.SetVar(IList_Item) = mArray(idx)

End Property
'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IList_NewEnum() As stdole.IUnknown
    
    '配列からCollection生成
    Set mCol = Core.ArrayToCollection(mArray)
    
    Set IList_NewEnum = mCol.[_NewEnum]

End Function
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

    Dim i As Long
    
    If idx < 0 Or idx >= mCount Then
        Err.Raise 512 + 1, TypeName(Me) & ".RemoveAt", "IndexOutOfRangeException"
    End If

    For i = idx To mCount - 1 - 1
        Core.SetVar(mArray(i)) = mArray(i + 1)
    Next

    mCount = mCount - 1
    
    If mCount = 0 Then
        mArray = Array()
    Else
        ReDim Preserve mArray(0 To mCount - 1)
    End If

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
Dim PL As PairLoggerInfo
Set PL = PairLoggerInfo.CreateObject(TypeName(Me) & ".IList_Sort")

    If CP Is Nothing Then
        'Interfaceも普通のクラスなのでDefault比較として利用
        Set CP = New IComparer
    End If

    QuickSort mArray, LBound(mArray), UBound(mArray), CP

End Sub
Private Sub QuickSort(ByRef ary As Variant, ByVal Min As Long, ByVal Max As Long, ByRef CP As IComparer)
    
    Dim i As Long
    Dim j As Long
    
    Dim base As Variant
    Dim work As Variant
    
    Core.SetVar(base) = ary((Min + Max) \ 2)
    
    i = Min
    j = Max
    
    Do
        
        Do While CP.Compare(ary(i), base) < 0
            i = i + 1
        Loop
        
        Do While CP.Compare(ary(j), base) > 0
            j = j - 1
        Loop
        
        If i >= j Then
            Exit Do
        End If
        
        '入れ替え
        Core.SetVar(work) = ary(i)
        Core.SetVar(ary(i)) = ary(j)
        Core.SetVar(ary(j)) = work
        
        i = i + 1
        j = j - 1
    
    Loop
    
    If (Min < i - 1) Then
        Call QuickSort(ary, Min, i - 1, CP)
    End If
    
    If (Max > j + 1) Then
        Call QuickSort(ary, j + 1, Max, CP)
    End If
 
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
    IList_ToArray = mArray
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
    '配列からCollection生成
    Set IList_ToCollection = Core.ArrayToCollection(mArray)
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
    Dim SB As IStringBuilder
    Set SB = StringBuilder.CreateObject
    
    For Each v In mArray
1
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
                    
                        SB.Append """" & Escape(v) & """"
                    
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
Private Function Escape(ByVal v As String) As String

    Dim strChr As String
    Dim SB As IStringBuilder
    Dim i As Long
    
    Set SB = StringBuilder.CreateObject
    
    For i = 1 To Len(v)
        
        strChr = Mid$(v, i, 1)
        
        Select Case AscW(strChr)
            Case &H0 To &H7, &HB, &HE To &H1F, &H7F
                SB.Append "\u" & Hex$(AscW(strChr))
            Case &H8
                SB.Append "\b"
            Case &H9
                SB.Append "\t"
            Case &HA
                SB.Append "\n"
            Case &HC
                SB.Append "\f"
            Case &HD
                SB.Append "\r"
            Case &H22
                SB.Append "\"""
            Case &H2F
                SB.Append "\/"
            Case &H5C
                SB.Append "\\"
            Case Else
                SB.Append strChr
        End Select

    Next
    
    Escape = SB.ToString

End Function
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
    ICursor_Eof = UBound(mArray) < mIndex
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
    If ICursor_Eof Then
        ICursor_Item = mArray(UBound(mArray))
    Else
        ICursor_Item = mArray(mIndex)
    End If
End Property
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
Private Sub ICursor_MoveFirst()
    mIndex = LBound(mArray)
End Sub
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
Private Sub ICursor_MoveNext()
    mIndex = mIndex + 1
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
