VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
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
' 文字列連結クラス
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor

Private mstrBuf() As String
Private mlngCount As Long
Private mlngLength As Long
Private Const C_INIT_COUNT As Long = 25
Private mlngInitCount As Long
'------------------------------------------------------
' 初期化
'------------------------------------------------------
Private Sub Class_Initialize()
    mlngInitCount = C_INIT_COUNT
    Me.Clear
End Sub
'------------------------------------------------------
' 終了
'------------------------------------------------------
Private Sub Class_Terminate()
    Erase mstrBuf
End Sub
'------------------------------------------------------
' コンストラクタ　初期配列サイズの変更
'------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    If Args.Count = 1 Then
        mlngInitCount = Args(1)
        Me.Clear
    End If
    
    Set IConstructor_Instancing = Me
    
End Function
'------------------------------------------------------
' クリア
'------------------------------------------------------
Public Sub Clear()
    
    Erase mstrBuf
    ReDim Preserve mstrBuf(0 To mlngInitCount)
    
    mlngCount = 0
    mlngLength = 0

End Sub
'------------------------------------------------------
' 追加
'------------------------------------------------------
Function Append(ByVal s As String, Optional ByVal quotation As Boolean = False)

    '配列にセット
    If mlngCount > UBound(mstrBuf) Then
        ReDim Preserve mstrBuf(0 To mlngCount)
    End If
    
    If quotation Then
        mstrBuf(mlngCount) = """" & s & """"
    Else
        mstrBuf(mlngCount) = s
    End If
    
    '文字数をカウント
    mlngLength = mlngLength + Len(mstrBuf(mlngCount))
    
    '要素数＋１
    mlngCount = mlngCount + 1

End Function
'------------------------------------------------------
' 文字列変換
'------------------------------------------------------
Function ToString()

    Call resize
    ToString = Join(mstrBuf, "")

End Function
'------------------------------------------------------
' 文字列変換(JOIN)
'------------------------------------------------------
Function ToJoin(ByVal strDelimiter As String)
    
    Call resize
    ToJoin = Join(mstrBuf, strDelimiter)

End Function
'------------------------------------------------------
' 配列戻し
'------------------------------------------------------
Function ToArray() As String()
    
    Call resize
    ToArray = mstrBuf()

End Function
'------------------------------------------------------
' 文字列変換前のリサイズ
'------------------------------------------------------
Private Sub resize()
    
    Select Case mlngCount
        Case Is <= 0
            ReDim Preserve mstrBuf(0 To 0)
        Case Is < mlngInitCount
            ReDim Preserve mstrBuf(0 To mlngCount - 1)
    End Select

End Sub
'------------------------------------------------------
' 文字数
'------------------------------------------------------
Public Property Get Length() As Long
    Length = mlngLength
End Property
'------------------------------------------------------
' 置換
'------------------------------------------------------
Public Sub Replace(ByVal strFind As String, ByVal strReplace As String)

    Dim i As Long
    If mlngCount > 0 Then
        For i = 0 To mlngCount - 1
            mstrBuf(i) = VBA.Replace(mstrBuf(i), strFind, strReplace)
        Next
    End If
    
End Sub

