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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。配列関係のメソッドがあります。
'-----------------------------------------------------------------------------------------------------
' ArrayList クラス
'-----------------------------------------------------------------------------------------------------
' TODO : ArrayList クラス Collection をベースにしているが、添字は0からなので注意
'
' ArrayList.IsArray(var) (Static)
'
' Dim list As IList
' Set list = New ArrayList
'
' Dim list As IList
' Set list = Constructor(New ArrayList, コレクション)
'
' Dim list As IList
' Set list = Constructor(New ArrayList, １次元配列)
'
'-----------------------------------------------------------------------------------------------------
' NOTE : ToString で JSON Array を出力します。
'        https://www.json.org/json-ja.html
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
Public Function GetIList() As IList
    Set GetIList = Me
End Function
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - Collection  : 1D Only
' Arg2 - Array       : 1D Only
'--------------------------------------------------------------
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
'--------------------------------------------------------------
' Static メソッド
'--------------------------------------------------------------
' 空配列対応 IsArray
'---------------------------------------------------------------------
' Dim a() As Variant --> Error 9 インデックスが有効範囲にありません。
' a = Array() --> Ubound が -1
'---------------------------------------------------------------------
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
'-------------------------------------------
' Static メソッド
'-------------------------------------------
' エラー検知付き Ubound
'-------------------------------------------
Public Function TryUBound(ary As Variant, lngIdx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    lngIdx = UBound(ary, dimension)

    TryUBound = True
    
    Exit Function
e:
    TryUBound = False

End Function
'-------------------------------------------
' Static メソッド
'-------------------------------------------
' エラー検知付き Lbound
'-------------------------------------------
Public Function TryLBound(ary As Variant, lngIdx As Long, Optional ByVal dimension As Long = 1) As Boolean
 
    On Error GoTo e
    
    lngIdx = LBound(ary, dimension)

    TryLBound = True
    
    Exit Function
e:
    TryLBound = False

End Function
''-------------------------------------------
'' １～２次元配列からArrayListに変換
''-------------------------------------------
'Public Function ParseArray(ary As Variant, Optional obj As Object = Nothing, Optional Fields As Variant) As Object
'
'    Dim lngVal As Long
'    Dim lineStart As Long
'    Dim lineStop As Long
'    Dim fieldStart As Long
'    Dim fieldStop As Long
'
'    '2次元フラグ
'    Dim bln2d As Boolean
'
'    bln2d = True
'
'    '1次元目
'    If TryLBound(ary, lngVal) Then
'        lineStart = lngVal
'        If TryUBound(ary, lngVal) Then
'            lineStop = lngVal
'        Else
'            Message.Throw 1, Me, "ParseArray", "not array"
'        End If
'    Else
'        Message.Throw 1, Me, "ParseArray", "not array"
'    End If
'
'    '2次元目
'    If TryLBound(ary, lngVal, 2) Then
'        fieldStart = lngVal
'        If TryUBound(ary, lngVal, 2) Then
'            fieldStop = lngVal
'        Else
'            '2次元ではない
'            bln2d = False
'        End If
'    Else
'        '2次元ではない
'        bln2d = False
'    End If
'
'    Dim i As Long
'    Dim j As Long
'
'    Dim obj1 As Object
'    Dim obj2 As Object
'
'    '2次元配列の場合
'    If bln2d Then
'
'        '引数によりインスタンス化するオブジェクトを判定
'        Select Case True
'            Case obj Is Nothing
'                Set obj1 = GetNewInstance(New ArrayList)
'            Case TypeOf obj Is Collection
'                Set obj1 = GetNewInstance(obj(1))
'            Case Else
'                Set obj1 = GetNewInstance(obj)
'        End Select
'
'        For i = lineStart To lineStop
'
'            '引数によりインスタンス化するオブジェクトを判定
'            Select Case True
'                Case obj Is Nothing
'                    Set obj2 = GetNewInstance(New ArrayList)
'                Case TypeOf obj Is Collection
'                    Set obj2 = GetNewInstance(obj(2))
'                Case Else
'                    Set obj2 = GetNewInstance(obj)
'            End Select
'
'            For j = fieldStart To fieldStop
'
'                If IsMissing(Fields) Then
'                    ToIDictionaryAccessor(obj2).SetProperty "", ary(i, j)
'                Else
'                    ToIDictionaryAccessor(obj2).SetProperty Fields(j), ary(i, j)
'                End If
'
'            Next
'
'            ToIDictionaryAccessor(obj1).SetProperty "", obj2
'
'        Next
'
'    '1次元配列の場合
'    Else
'
'        '引数によりインスタンス化するオブジェクトを判定
'        Select Case True
'            Case obj Is Nothing
'                Set obj1 = GetNewInstance(New ArrayList)
'            Case TypeOf obj Is Collection
'                Set obj1 = GetNewInstance(obj(1))
'            Case Else
'                Set obj1 = GetNewInstance(obj)
'        End Select
'
'        For i = lineStart To lineStop
'
'            If IsMissing(Fields) Then
'                ToIDictionaryAccessor(obj1).SetProperty "", ary(i)
'            Else
'                ToIDictionaryAccessor(obj1).SetProperty Fields(i), ary(i)
'            End If
'
'        Next
'
'    End If
'
'    Set ParseArray = obj1
'
'End Function
'-------------------------------------------
' Listbox（Combobox）からArrayList変換
' Tagにフィールド情報（カンマ区切り）が必須。
'-------------------------------------------
Public Function ParseArrayFromListbox(objList As Object, Optional obj As Object = Nothing) As Object

    Dim ary As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim i As Long
    Dim j As Long
    
    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Message.Throw 1, Me, "ParseArrayFromListbox", "never set tag in 'fields string'"
    End If
    
    '配列に変換
    Fields = Split(objList.Tag, ",")
    
    '1次元目はArrayList固定
    Set ary = New ArrayList
        
    For i = 0 To objList.ListCount - 1
        
        '引数によりインスタンス化するオブジェクトを判定
        Select Case True
            Case obj Is Nothing
                Set objRec = GetNewInstance(New Dictionary)
            Case Else
                Set objRec = GetNewInstance(obj)
        End Select
        
        For j = 0 To objList.ColumnCount - 1
        
            '設定されているフィールドのみコピーする。
            If Fields(j) <> "" Then
                ToIDictionaryAccessor(objRec).SetProperty Fields(j), objList.List(i, j)
            End If
        Next
            
        ary.Add objRec
    
    Next

    Set ParseArrayFromListbox = ary

End Function
'-------------------------------------------
' ArrayListの内容をListbox(Combobox)にコピーする。
' Tagにフィールド情報（カンマ区切り）が必須。
'-------------------------------------------
Public Sub CopyToListbox(objList As Object)

    Dim ary As IList
    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant

    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Message.Throw 1, Me, "CopyToListbox", "never set tag in 'fields string'"
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
                objList.List(i, j) = ToIDictionaryAccessor(objRec).GetProperty(v)
            End If
            j = j + 1
        Next
        i = i + 1
    Next

End Sub

Function GetNewInstance(obj As INewInstance) As Object
    Set GetNewInstance = obj.NewInstance
End Function
'-------------------------------------------------
' ToIDictionaryAccessorへキャスト
'-------------------------------------------------
Private Function ToIDictionaryAccessor(obj As Object) As IDictionaryAccessor
    Set ToIDictionaryAccessor = obj
End Function
'--------------------------------------------------------------
' ArrayList から全ての項目を削除する。
'--------------------------------------------------------------
Private Sub IList_Clear()

    Set mCol = New Collection
    mIndex = 0
    
End Sub
'--------------------------------------------------------------
' ArrayList に項目を追加する。
'--------------------------------------------------------------
Private Sub IList_Add(obj As Variant)
    
    mCol.Add obj

End Sub
'--------------------------------------------------------------
' ArrayList の指定位置に項目を追加する。
'--------------------------------------------------------------
Private Sub IList_Insert(Idx As Long, v As Variant)

    If Idx < 0 Or Idx >= mCol.Count Then
        Message.Throw 1, Me, "Insert", "IndexOutOfRangeException"
    End If
    
    mCol.Add v, , Idx + 1

End Sub
'--------------------------------------------------------------
' For Each 用
'--------------------------------------------------------------
Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = mCol.[_NewEnum]
End Function
'--------------------------------------------------------------
' ArrayList 要素数を返却
'--------------------------------------------------------------
Private Property Get IList_Count() As Long
    IList_Count = mCol.Count
End Property
'--------------------------------------------------------------
' ArrayList の指定位置の項目を取得
'--------------------------------------------------------------
Private Property Get IList_Item(Idx As Long) As Variant
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Message.Throw 1, Me, "Item", "IndexOutOfRangeException"
    End If

    If VBA.IsObject(mCol.Item(Idx + 1)) Then
        Set IList_Item = mCol.Item(Idx + 1)
    Else
        IList_Item = mCol.Item(Idx + 1)
    End If
    
End Property
'--------------------------------------------------------------
' ArrayList の指定位置の項目を削除
'--------------------------------------------------------------
Private Sub IList_RemoveAt(Idx As Long)
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Message.Throw 1, Me, "RemoveAt", "IndexOutOfRangeException"
    End If
    
    mCol.Remove Idx + 1
End Sub
'--------------------------------------------------------------
' ArrayList のソート
'--------------------------------------------------------------
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
'--------------------------------------------------------------
' ArrayList を１～２次元配列に変換
'--------------------------------------------------------------
Private Function IList_ToArray() As Variant

    Dim v As Variant
    Dim s As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    Dim j As Long
    Dim result As Variant
    Dim List As IList
    
    result = Array()
    
    lngRow = mCol.Count
    lngCol = 0
    
    Set List = Me
    
    'ArrayList/ArrayList の２次元を想定
    For Each v In List
        If TypeName(v) = "ArrayList" Then
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
        
            For i = 0 To List.Count - 1
                If IsObject(List(i)) Then
                    Set result(i) = List(i)
                Else
                    Let result(i) = List(i)
                End If
            Next
    
            IList_ToArray = result
        
        '2Dの場合
        Case lngRow > 0 And lngCol > 0
        
            ReDim result(0 To lngRow - 1, 0 To lngCol - 1)
        
            For i = 0 To List.Count - 1
                For j = 0 To List(i).Count - 1
                    If IsObject(List(i)(j)) Then
                        Set result(i, j) = List(i)(j)
                    Else
                        Let result(i, j) = List(i)(j)
                    End If
                Next
            Next
    
            IList_ToArray = result
    End Select
    
End Function
'--------------------------------------------------------------
' ArrayList をCollectionに変換
'--------------------------------------------------------------
Private Function IList_ToCollection() As Collection
    
    Set IList_ToCollection = mCol

End Function
'--------------------------------------------------------------
' ArrayList をJSONに変換
'--------------------------------------------------------------
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
                        Message.Throw 1, Me, "ToString", "not cast type ""{0}""", TypeName(v)
        
                End Select
        End Select
        
    Next

    IList_ToString = "[" & SB.ToJoin(", ") & "]"

End Function
'--------------------------------------------------------------
' 制御文字変換
'---------------------------------------------------------------------
Private Function Escape(ByVal v As String) As String

    Escape = """" & Convert.Escape(v) & """"

End Function
'Unicode
Private Function Unicode(ByVal strBuf As String) As Long
    Dim bytBuf() As Byte
    
    If Len(strBuf) <> 0 Then
        bytBuf = strBuf
        Unicode = CLng(bytBuf(1)) * &H100 + bytBuf(0)
    End If
End Function
'--------------------------------------------------------------
' ArrayList の ICursor インターフェースを取得
'--------------------------------------------------------------
Private Property Get IList_GetCursor() As ICursor
    Set IList_GetCursor = Me
End Property
'--------------------------------------------------------------
' ICursor の Eof プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mIndex >= mCol.Count
End Property
'--------------------------------------------------------------
' ICursor の Item プロパティを取得
'--------------------------------------------------------------
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
'--------------------------------------------------------------
' ICursor の カーソル位置を初期化
'--------------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = 0
End Sub
'--------------------------------------------------------------
' ICursor の カーソル位置を次に移動
'--------------------------------------------------------------
Private Sub ICursor_MoveNext()

    mIndex = mIndex + 1
    If mIndex > mCol.Count Then
        mIndex = mCol.Count
    End If

End Sub
'--------------------------------------------------------------
' インスタンス生成 I/F
'--------------------------------------------------------------
Private Function INewInstance_NewInstance() As Object

    Dim lst As IList
    
    Set lst = New ArrayList
    
    Set INewInstance_NewInstance = lst

End Function
