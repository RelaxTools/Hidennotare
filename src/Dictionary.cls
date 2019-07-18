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
' Dictionary ラッパークラス
'-----------------------------------------------------------------------------------------------------
' TODO : 本家と異なる点。For Each を使う場合には XX.Keys を省略できない。
'        IDictionary I/F を使用する必要がある。
'
' Dim dic As IDictionary
' Set dic = New Dictionary
'
'-----------------------------------------------------------------------------------------------------
' ToString すると JSON Object を出力します。
' https://www.json.org/json-ja.html
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements INewInstance
Implements IDictionaryAccessor

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
'  IDictionary I/F 取得
'--------------------------------------------------------------
Public Function GetIDictionary() As IDictionary
    Set GetIDictionary = Me
End Function
'-------------------------------------------
' Listbox（Combobox）からDictionary変換
' Tagにフィールド情報（カンマ区切り）が必須。
'-------------------------------------------
Public Function ParseFromListbox(objList As Object, KeyName As String, Optional obj As Object = Nothing) As IDictionary

    Dim ary As IDictionary
    Dim objRec As Object
    Dim Fields As Variant
    Dim i As Long
    Dim j As Long
    Dim Key As Variant
    
    'Tagにフィールド情報が無い場合
    If objList.Tag = "" Then
        Message.Throw 1, Me, "ParseFromListbox", "never set tag in 'fields string'"
    End If
    
    '配列に変換
    Fields = Split(objList.Tag, ",")
    
    '1次元目はArrayList固定
    Set ary = New Dictionary
        
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
                
                If KeyName = Fields(j) Then
                    Key = objList.List(i, j)
                End If
            
            End If
        Next
            
        ary.Add Key, objRec
    
    Next

    Set ParseFromListbox = ary

End Function
'-------------------------------------------
' Dictionaryの内容をListbox(Combobox)にコピーする。
' Tagにフィールド情報（カンマ区切り）が必須。
'-------------------------------------------
Public Sub CopyToListbox(objList As Object)

    Dim ary As IDictionary
    Dim objRec As Object
    Dim Fields As Variant
    Dim v As Variant
    Dim Key As Variant

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
    For Each Key In ary.Keys
    
        Set objRec = ary.Item(Key)
    
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
'-------------------------------------------
' ListObject（テーブル機能）からDictionary変換
' テーブルの列ヘッダの情報が必須。
'-------------------------------------------
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
                Set objRec = GetNewInstance(New Dictionary)
            Case Else
                Set objRec = GetNewInstance(obj)
        End Select
        
        For j = 1 To LO.HeaderRowRange.columns.Count
            ToIDictionaryAccessor(objRec).SetProperty LO.HeaderRowRange(, j).Value, LO.DataBodyRange.Cells(i, j).Value
        Next
    
        'Keyフィールド取得
        j = LO.ListColumns(KeyField).Index
        Key = LO.DataBodyRange.Cells(i, j).Value
        
        If lst.ContainsKey(Key) Then
            Message.Throw 1, Me, "ParseFromListObject", "Duplicate Key"
        End If
        
        lst.Add Key, objRec
    
    Next

    Set ParseFromListObject = lst

End Function
'-------------------------------------------
' Dictionaryの内容をListObject(テーブル機能)にコピーする。
' テーブルの列ヘッダの情報が必須。
'-------------------------------------------
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
        For j = 1 To LO.HeaderRowRange.columns.Count

            LO.DataBodyRange.Cells(i, j).Value = ToIDictionaryAccessor(obj).GetProperty(LO.HeaderRowRange(, j).Value)
        
        Next
        i = i + 1
    Next

End Sub
'-------------------------------------------------
' NewInstance
'-------------------------------------------------
Private Function GetNewInstance(obj As INewInstance) As Object
    Set GetNewInstance = obj.NewInstance
End Function
'-------------------------------------------------
' ToIDictionaryAccessorへキャスト
'-------------------------------------------------
Private Function ToIDictionaryAccessor(obj As Object) As IDictionaryAccessor
    Set ToIDictionaryAccessor = obj
End Function
'--------------------------------------------------------------
'  項目の追加
'--------------------------------------------------------------
Private Sub IDictionary_Add(KeyValue As Variant, obj As Variant)
    mDic.Add KeyValue, obj
End Sub
'--------------------------------------------------------------
'  項目の存在チェック
'--------------------------------------------------------------
Private Function IDictionary_ContainsKey(KeyValue As Variant) As Boolean
    IDictionary_ContainsKey = mDic.Exists(KeyValue)
End Function
'--------------------------------------------------------------
'  キーリスト取得
'--------------------------------------------------------------
Private Function IDictionary_Keys() As Variant
    IDictionary_Keys = mDic.Keys
End Function
'--------------------------------------------------------------
'  値リスト取得
'--------------------------------------------------------------
Private Property Get IDictionary_Count() As Long
    IDictionary_Count = mDic.Count
End Property
'--------------------------------------------------------------
'  値取得
'--------------------------------------------------------------
Private Property Get IDictionary_Item(KeyValue As Variant) As Variant
    If VBA.IsObject(mDic.Item(KeyValue)) Then
        Set IDictionary_Item = mDic.Item(KeyValue)
    Else
        IDictionary_Item = mDic.Item(KeyValue)
    End If
End Property
'--------------------------------------------------------------
'  値設定
'--------------------------------------------------------------
Private Property Let IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  値設定(オブジェクト)
'--------------------------------------------------------------
Private Property Set IDictionary_Item(KeyValue As Variant, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  キー設定
'--------------------------------------------------------------
Private Property Let IDictionary_Key(OldKey As Variant, NewKey As Variant)
    mDic.Key(OldKey) = NewKey
End Property
'--------------------------------------------------------------
'  削除
'--------------------------------------------------------------
Private Sub IDictionary_Remove(KeyValue As Variant)
    mDic.Remove KeyValue
End Sub
'--------------------------------------------------------------
'  全アイテム削除
'--------------------------------------------------------------
Private Sub IDictionary_Clear()
    mDic.RemoveAll
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  シリアライズ化(JSON)
'--------------------------------------------------------------
Private Function IDictionary_ToString() As String
    IDictionary_ToString = ToString()
End Function
'--------------------------------------------------------------
'  シリアライズ化(JSON)
'--------------------------------------------------------------
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
                        Message.Throw 1, Me, "ToString", "not cast type ""{0}""", TypeName(v)
        
                End Select

        End Select
    
    Next

    ToString = "{" & SB.ToJoin(", ") & "}"

End Function

' 制御文字変換
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
'  ICursor の ICursor インターフェースを取得
'--------------------------------------------------------------
Private Property Get IDictionary_GetCursor() As ICursor
    Set IDictionary_GetCursor = Me
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
    Dim v As Variant
    
    v = IDictionary_Keys
    
    If mIndex > mDic.Count - 1 Then
        ICursor_Item = v(mDic.Count - 1)
    Else
        ICursor_Item = v(mIndex)
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
    Set d = New Dictionary
    
    Set INewInstance_NewInstance = d

End Function
'--------------------------------------------------------------
' IDictionaryAccessor I/F
'--------------------------------------------------------------
Private Sub IDictionaryAccessor_SetProperty(prop As Variant, Val As Variant)
    
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
Private Function IDictionaryAccessor_GetProperty(prop As Variant) As Variant
    If IsObject(Me.GetIDictionary.Item(prop)) Then
        Set IDictionaryAccessor_GetProperty = Me.GetIDictionary.Item(prop)
    Else
        Let IDictionaryAccessor_GetProperty = Me.GetIDictionary.Item(prop)
    End If
End Function
