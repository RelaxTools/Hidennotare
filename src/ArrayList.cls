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
' [Hecatoncheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hecatoncheir
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
'  ArrayList クラス
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor
Implements IList

Private mCol As Collection
Private v() As Variant
Private mIndex As Long
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mCol = New Collection
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--------------------------------------------------------------
'  コンストラクタ
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
'---------------------------------------------------------------------------------------------------
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
'--------------------------------------------------------------
'  ArrayList から全ての項目を削除する。
'--------------------------------------------------------------
Private Sub IList_Clear()

    Set mCol = New Collection
    mIndex = 0
    
End Sub
'--------------------------------------------------------------
'  ArrayList に項目を追加する。
'--------------------------------------------------------------
Private Sub IList_Add(obj As Variant)
    
    mCol.Add obj

End Sub
'--------------------------------------------------------------
'  ArrayList の指定位置に項目を追加する。
'--------------------------------------------------------------
Private Sub IList_Insert(Idx As Long, v As Variant)

    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".Insert", "IndexOutOfRangeException"
    End If
    
    mCol.Add v, , Idx + 1

End Sub
'--------------------------------------------------------------
'  For Each 用
'--------------------------------------------------------------
Private Function IList_NewEnum() As stdole.IUnknown
    Set IList_NewEnum = mCol.[_NewEnum]
End Function
'--------------------------------------------------------------
'  ArrayList 要素数を返却
'--------------------------------------------------------------
Private Property Get IList_Count() As Long
    IList_Count = mCol.Count
End Property
'--------------------------------------------------------------
'  ArrayList の指定位置の項目を取得
'--------------------------------------------------------------
Private Property Get IList_Item(Idx As Long) As Variant
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".Item", "IndexOutOfRangeException"
    End If
        
    If VBA.IsObject(mCol.Item(Idx + 1)) Then
        Set IList_Item = mCol.Item(Idx + 1)
    Else
        IList_Item = mCol.Item(Idx + 1)
    End If
    
End Property
'--------------------------------------------------------------
'  ArrayList の指定位置の項目を削除
'--------------------------------------------------------------
Private Sub IList_RemoveAt(Idx As Long)
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".RemoveAt", "IndexOutOfRangeException"
    End If
    
    mCol.Remove Idx + 1
End Sub
'--------------------------------------------------------------
'  ArrayList のソート
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
'  ArrayList を１次元配列に変換
'--------------------------------------------------------------
Private Function IList_ToArray() As Variant

    Dim c As Variant
    Dim i As Long
    
    ReDim v(0 To mCol.Count)
    
    For i = 0 To mCol.Count - 1

        If IsObject(mCol(i + 1)) Then
            Set v(i) = mCol(i + 1)
        Else
            Let v(i) = mCol(i + 1)
        End If
        
    Next

    IList_ToArray = v
    
End Function
'--------------------------------------------------------------
'  ArrayList をCollectionに変換
'--------------------------------------------------------------
Private Function IList_ToCollection() As Collection
    
    Set IList_ToCollection = mCol

End Function
'--------------------------------------------------------------
'  ArrayList をJSONに変換
'--------------------------------------------------------------
Private Function IList_ToString() As String

    Dim v As Variant
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    
    For Each v In mCol
    
        Select Case True
            Case IsObject(v), VarType(v) = vbDataObject
                
                sb.Append v.ToString
            
            Case IsEmpty(v)
                
                sb.Append "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        sb.Append LCase(CStr(v))
        
                    Case vbString
                    
                        sb.Append IList_Escape(v)
                    
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal ', vbLongLong
        
                        sb.Append v
        
                    Case Else
        
                        'エラー
                        Err.Raise vbObjectError + 1, TypeName(Me) & ".ToString", Message.PlaceHolder("not cast type ""{0}""", TypeName(v))
        
                End Select
        End Select
        
    Next

    IList_ToString = "[" & sb.ToJoin(", ") & "]"

End Function
'---------------------------------------------------------------------
' 制御文字変換
'---------------------------------------------------------------------
Private Function IList_Escape(ByVal v As String) As String

    Dim strChr As String
    Dim sb As StringBuilder
    Dim i As Long
    
    Set sb = New StringBuilder
    
    For i = 1 To Len(v)
        
        strChr = Mid$(v, i, 1)
        
        Select Case Unicode(strChr)
            Case &H0 To &H7, &HB, &HE To &H1F, &H7F
                sb.Append "\u" & Hex$(Unicode(strChr))
            Case &H8
                sb.Append "\b"
            Case &H9
                sb.Append "\t"
            Case &HA
                sb.Append "\n"
            Case &HC
                sb.Append "\f"
            Case &HD
                sb.Append "\r"
            Case &H22
                sb.Append "\"""
            Case &H2F
                sb.Append "\/"
            Case &H5C
                sb.Append "\\"
            Case Else
                sb.Append strChr
        End Select

    Next
    
    '本当にJSONするならここでUTF8変換
    IList_Escape = """" & sb.ToString & """"

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
'  ArrayList の ICursor インターフェースを取得
'--------------------------------------------------------------
Private Property Get IList_GetCursor() As ICursor
    Set IList_GetCursor = Me
End Property
'--------------------------------------------------------------
'  ICursor の Eof プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mIndex >= mCol.Count
End Property
'--------------------------------------------------------------
'  ICursor の Item プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    
    If mIndex >= mCol.Count Then
        ICursor_Item = mCol(mCol.Count)
    Else
        ICursor_Item = mCol(mIndex + 1)
    End If
    
End Property
'--------------------------------------------------------------
'  ICursor の PreviousItem プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    
    If mIndex <= 0 Then
        Err.Raise vbObjectError + 512 + 1, , "IndexOutOfRangeException"
    Else
        ICursor_PreviousItem = mCol(mIndex)
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
    If mIndex > mCol.Count Then
        mIndex = mCol.Count
    End If

End Sub

