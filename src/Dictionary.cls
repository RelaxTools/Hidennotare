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
' Dictionary ラッパークラス
'-----------------------------------------------------------------------------------------------------
' TODO : 本家と異なる点。For Each を使う場合には XX.Keys を省略できない。
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IDictionary
Implements INewInstance
Implements IJSONObject

Private mDic As Object
Private mIndex As Long
'--------------------------------------------------------------
'  IDictionary I/F 取得
'--------------------------------------------------------------
Public Function GetIDictionary() As IDictionary
    Set GetIDictionary = Me
End Function
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
Private Function IDictionary_Add(ByVal KeyValue As String, obj As Variant)
    mDic.Add KeyValue, obj
End Function
'--------------------------------------------------------------
'  項目の存在チェック
'--------------------------------------------------------------
Private Function IDictionary_ContainsKey(ByVal KeyValue As String) As Boolean
    IDictionary_ContainsKey = mDic.Exists(KeyValue)
End Function
'--------------------------------------------------------------
'  キーリスト取得
'--------------------------------------------------------------
Private Property Get IDictionary_Keys() As Variant
    IDictionary_Keys = mDic.Keys
End Property
'--------------------------------------------------------------
'  値リスト取得
'--------------------------------------------------------------
Private Property Get IDictionary_Count() As Long
    IDictionary_Count = mDic.Count
End Property
'--------------------------------------------------------------
'  値取得
'--------------------------------------------------------------
Private Property Get IDictionary_Item(ByVal KeyValue As String) As Variant
    If VBA.IsObject(mDic.Item(KeyValue)) Then
        Set IDictionary_Item = mDic.Item(KeyValue)
    Else
        IDictionary_Item = mDic.Item(KeyValue)
    End If
End Property
'--------------------------------------------------------------
'  値設定
'--------------------------------------------------------------
Private Property Let IDictionary_Item(ByVal KeyValue As String, NewItem As Variant)
    mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  値設定(オブジェクト)
'--------------------------------------------------------------
Private Property Set IDictionary_Item(ByVal KeyValue As String, NewItem As Variant)
    Set mDic.Item(KeyValue) = NewItem
End Property
'--------------------------------------------------------------
'  キー設定
'--------------------------------------------------------------
Private Property Let IDictionary_Key(ByVal OldKey As String, NewKey As Variant)
    mDic.Key(OldKey) = NewKey
End Property
'--------------------------------------------------------------
'  削除
'--------------------------------------------------------------
Private Sub IDictionary_Remove(ByVal KeyValue As String)
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
Private Function ToString(Optional ByVal LQuot As String = "{", Optional ByVal RQuot As String = "}") As String

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
            Case IsObject(v), VarType(v) = vbDataObject
                
                SB.Append """" & KeyValue & """:" & v.ToString
            
            Case IsEmpty(v)
                
                SB.Append """" & KeyValue & """:" & "null"
            
            Case Else
                Select Case VarType(v)
                    
                    Case vbBoolean
        
                        SB.Append """" & KeyValue & """:" & LCase(CStr(v))
        
                    Case vbString
                    
                        SB.Append """" & KeyValue & """:" & Escape(v)
                    
                    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal ', vbLongLong
        
                        SB.Append """" & KeyValue & """:" & v
        
                    Case Else
        
                        'エラー
                        Message.Throw 1, Me, "ToString", "not cast type ""{0}""", TypeName(v)
        
                End Select

        End Select
    
    Next

    ToString = LQuot & SB.ToJoin(", ") & RQuot

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
' JSON Object I/F
'--------------------------------------------------------------
Private Sub IJSONObject_SetProperty(prop As Variant, Val As Variant)
    
    Me.GetIDictionary.Add prop, Val
    
End Sub

