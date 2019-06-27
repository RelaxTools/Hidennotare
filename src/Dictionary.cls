VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'  Dictionary ラッパークラス (参照設定はしたくないけどインテリセンスしたいわがままクラス）
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor

Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum

Private mDic As Object
Private mIndex As Long
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mDic = CreateObject("Scripting.Dictionary")
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set mDic = Nothing
End Sub
'--------------------------------------------------------------
'  項目の追加
'--------------------------------------------------------------
Public Function Add(ByVal KeyValue As String, obj As Variant)
    mDic.Add KeyValue, obj
End Function
'--------------------------------------------------------------
'  項目の存在チェック
'--------------------------------------------------------------
Public Function Exists(ByVal KeyValue As String) As Boolean
    Exists = mDic.Exists(KeyValue)
End Function
'--------------------------------------------------------------
'  キーリスト取得
'--------------------------------------------------------------
Public Function keys() As Variant
    keys = mDic.keys
End Function
'--------------------------------------------------------------
'  ソート済みキーリスト取得
'--------------------------------------------------------------
Public Function SortedKeys(Optional ByRef CP As IComparer = Nothing) As Variant
    SortedKeys = Sort(CP)
End Function
'--------------------------------------------------------------
'  値リスト取得
'--------------------------------------------------------------
Public Function Items() As Variant
    Items = mDic.Items
End Function
Public Property Get CompareMode() As CompareMethod
    CompareMode = mDic.CompareMode
End Property
Public Property Let CompareMode(Mode As CompareMethod)
    mDic.CompareMode = Mode
End Property
Public Property Get Count() As Long
    Count = mDic.Count
End Property
'--------------------------------------------------------------
'  値取得
'--------------------------------------------------------------
Public Property Get Item(ByVal KeyValue As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(mDic.Item(KeyValue)) Then
        Set Item = mDic.Item(KeyValue)
    Else
        Item = mDic.Item(KeyValue)
    End If
End Property
'--------------------------------------------------------------
'  値設定
'--------------------------------------------------------------
Public Property Let Item(ByVal KeyValue As String, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  値設定(オブジェクト)
'--------------------------------------------------------------
Public Property Set Item(ByVal KeyValue As String, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  キー設定
'--------------------------------------------------------------
Public Property Let key(ByVal OldKey As String, NewKey As Variant)
    mDic.key(OldKey) = NewKey
End Property
'--------------------------------------------------------------
'  削除
'--------------------------------------------------------------
Public Sub Remove(ByVal KeyValue As String)
    mDic.Remove KeyValue
End Sub
'--------------------------------------------------------------
'  全アイテム削除
'--------------------------------------------------------------
Public Sub RemoveAll()
    
    mDic.RemoveAll
    mIndex = 0

End Sub
'--------------------------------------------------------------
'  キーのソート
'--------------------------------------------------------------
Private Function Sort(ByVal CP As IComparer) As Variant

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim wk As String
    Dim strKey As Variant
    
    If CP Is Nothing Then
        Set CP = New IComparer
    End If
    
    strKey = mDic.keys

    '要素数
    n = UBound(strKey) - LBound(strKey) + 1
    
    'ソート不要
    If n <= 1 Then
        GoTo e
    End If

    '挿入ソート
    For i = 1 To n - 1

        wk = strKey(i)
        
        Dim lngCmp As Long
        
        lngCmp = CP.Compare(strKey(i - 1), wk)

        If lngCmp > 0 Then

            j = i

            Do

                strKey(j) = strKey(j - 1)

                j = j - 1

                If j = 0 Then
                    Exit Do
                End If

                lngCmp = CP.Compare(strKey(j - 1), wk)
            
            Loop While lngCmp > 0
            
            strKey(j) = wk

        End If
    Next

e:
    Sort = strKey

End Function
'--------------------------------------------------------------
'  シリアライズ化(JSON)
'--------------------------------------------------------------
Public Function ToString() As String

    Dim Item  As Variant
    Dim KeyValue As Variant
    
    Dim v  As Variant
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    
    For Each KeyValue In mDic.keys
    
        If VBA.IsObject(mDic.Item(KeyValue)) Then
            Set v = mDic.Item(KeyValue)
        Else
            v = mDic.Item(KeyValue)
        End If
    
        Select Case True
            Case IsObject(v), VarType(v) = vbDataObject
                
                sb.Append """" & KeyValue & """:" & v.ToString
            
            Case IsEmpty(v)
                
                sb.Append """" & KeyValue & """:" & "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        sb.Append """" & KeyValue & """:" & LCase(CStr(v))
        
                    Case vbString
                    
                        sb.Append """" & KeyValue & """:" & Escape(v)
                    
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal ', vbLongLong
        
                        sb.Append """" & KeyValue & """:" & v
        
                    Case Else
        
                        'エラー
                        Message.Throw 1, Me, "ToString", "not cast type ""{0}""", TypeName(v)
        
                End Select

        End Select
    
    Next

    ToString = "{" & sb.ToJoin(", ") & "}"

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
Public Property Get GetCursor() As ICursor
    Set GetCursor = Me
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
    
    v = Me.SortedKeys
    
    If mIndex > mDic.Count - 1 Then
        ICursor_Item = v(mDic.Count - 1)
    Else
        ICursor_Item = v(mIndex)
    End If

End Property
'--------------------------------------------------------------
'  ICursor の PreviousItem プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    
    Dim v As Variant
    
    v = Me.SortedKeys
    
    If mIndex < 0 Then
        Message.Throw 1, Me, "ICursor_PreviousItem", "IndexOutOfRangeException"
    Else
        ICursor_PreviousItem = v(mIndex - 1)
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

