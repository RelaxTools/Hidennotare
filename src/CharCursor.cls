VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CharCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [hidennotare] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/hidennotare
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
' カーソル風文字列読み込みクラス
'-----------------------------------------------------------------------------------------------------
' Dim IC As ICursor
' Set IC = Constructor(New CharCursor, 文字列)
'
' Do Until IC.Eof
'
'     Debug.Print IC.Item '1文字
'
'     IC.MoveNext
' Loop
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

Private mBuf As String
Private mIndex As Long
Private Sub Class_Initialize()
    mIndex = 1
End Sub
'--------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------
' Arg1 - String   : String
'--------------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Select Case Args.Count
        Case 1
            mBuf = CStr(Args(1))
            ICursor_MoveFirst
        Case Else
            'エラー
            Exit Function
    End Select
    
    Set IConstructor_Instancing = Me
End Function
'--------------------------------------------------------------
' ICursor インターフェースを取得
'--------------------------------------------------------------
Public Property Get GetCursor() As ICursor
    Set GetCursor = Me
End Property
'------------------------------------------------------
' 終了判定
'------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = Len(mBuf) < mIndex
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = 1
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mIndex = mIndex + 1
End Sub
'------------------------------------------------------
' 文字取得(引数にて前後の文字を取得可）
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant

    Dim lngPos As Long
    
    If IsMissing(opt) Then
        opt = 0
    End If
    
    lngPos = mIndex + opt

    If lngPos < 1 Or lngPos > Len(mBuf) Then
        ICursor_Item = ""
    Else
        ICursor_Item = Mid$(mBuf, lngPos, 1)
    End If

End Property
