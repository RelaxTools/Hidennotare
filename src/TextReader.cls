VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
'>### TextReader クラス
'>
'>**Remarks**
'>
'>TextReaderクラス<br>
'>
'>**Dependent**
'>
'>* IConstructor クラス
'>* IReader クラス
'>* ICursor クラス
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements IReader
Implements ICursor

Private fp As Integer
Private mOpen As Boolean

'プロパティ保持
Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants

Private mBuf As Variant
Private mIndex As Long
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set IC = Constructor(New TextReader, FileName, [NewLineCode = NewLineCodeConstants.CRLF, Encode = EncodeConstants.CP932])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を設定|
'>|NewLineCode|省略可|NewLineCodeConstants|読み取り専用の設定。デフォルトCRLF|
'>|Encode|省略可|EncodeConstants|Applicationオブジェクト。デフォルトCP932|
'>
'>**Return Value**
'>
'>新しいインスタンス(TextReader)
'>
'>**Remarks**
'>
'>コンストラクタ<br>
'>With で生成した場合、End With でブックの Close をします。<br>
'>
'>**Example**
'>
'>1行ずつ読む場合(ICursorインターフェイス)
'>
'>```
'>Dim strFile As String
'>Dim IC As ICursor
'>
'>strFile = "C:\Test.xlsx"
'>
'>Set IC = Constructor(New TextReader, strFile)
'>
'>Do Until IC.Eof
'>
'>    Debug.Print IC.Item
'>    IC.MoveNext
'>Loop
'>```
'>
'>一気に読む場合(IReaderインターフェイス)
'>
'>```
'>Dim strFile As String
'>Dim IR As IReader
'>
'>strFile = "C:\Test.xlsx"
'>
'>Set IR = Constructor(New TextReader, strFile)
'>
'>Debug.Print IR.ReadAll '改行を指定した場合は配列、指定しなかった場合は文字列
'>```
'>
'>**See also**
'>
'>* None
'>
Private Function IConstructor_Instancing(Args As Collection) As Object

    If Args.Count = 0 Then
        Exit Function
    End If
    
    'Default
    mNewLineCode = NewLineCodeConstants.CRLF
    mEncode = EncodeConstants.CP932
    
    If Args.Count >= 1 Then
        mFileName = Args(1)
    End If
    
    If Args.Count >= 2 Then
        mNewLineCode = Args(2)
    End If
    
    If Args.Count >= 3 Then
        mEncode = Args(3)
    End If
    
    
    Set IConstructor_Instancing = Me

    
    If Not FileIO.FileExists(mFileName) Then
        Err.Raise 512 + 1, TypeName(Me) & ".IConstructor_Instancing", "file not found(" & mFileName & ")"
    End If

    'テキストファイルの一括読み込み
    mBuf = BulkRead()
    If IsArray(mBuf) Then
        ICursor_MoveFirst
    End If

End Function
'--------------------------------------------------------------
' テキストファイルの一括読み込み
'--------------------------------------------------------------
Private Function BulkRead() As Variant
    
    Dim buffer() As Byte
    Dim strBuf As String
    Dim lngSize As Long

    fp = FreeFile
    Open mFileName For Binary As fp
    mOpen = True
    
    lngSize = VBA.LOF(fp)
    
    If lngSize <> 0 Then

        ReDim buffer(0 To lngSize - 1)
    
        Get fp, , buffer()
        
        Select Case mEncode
            
            Case EncodeConstants.CP932
                
                strBuf = StrConv(buffer, vbUnicode)
            
            Case EncodeConstants.UTF8
            
                'BOMがある場合取り除く
                If StartWith(buffer, &HEF, &HBB, &HBF) Then
                    strBuf = Convert.FromUTF8(MidB$(buffer, 4))
                Else
                    strBuf = Convert.FromUTF8(buffer)
                End If
        
            Case EncodeConstants.UTF16LE
                
                'BOMがある場合取り除く
                If StartWith(buffer, &HFF, &HFE) Then
                    strBuf = MidB$(buffer, 3)
                Else
                    strBuf = buffer
                End If
        
        End Select
        
        Select Case mNewLineCode
            Case NewLineCodeConstants.CRLF
                
                BulkRead = Split(strBuf, vbCrLf)
            
            Case NewLineCodeConstants.LF
                
                BulkRead = Split(strBuf, vbLf)
        
            Case NewLineCodeConstants.None
                
                BulkRead = strBuf
        
        End Select

    Else
        BulkRead = Empty
    End If
    
    Close fp
    mOpen = False
        
End Function
'--------------------------------------------------------------
' BOMチェック
'--------------------------------------------------------------
Private Function StartWith(bytBuf, ParamArray Args() As Variant) As Boolean

    Dim v As Variant
    Dim i As Long
    
    i = 0
    For Each v In Args
            
        If i > UBound(bytBuf) Then
            StartWith = False
            Exit Function
        End If
        
        If Not bytBuf(i) = v Then
            StartWith = False
            Exit Function
        End If
        i = i + 1
    Next

    StartWith = True

End Function
'>---
'>#### 一括読み込み
'>
'>**Syntax**
'>
'>```
'>data = obj.ReadAll
'>```
'>
'>**Parameters**
'>
'>なし
'>
'>**Return Value**
'>
'>読み込んだデータを返却。<br>
'>NewLineCodeで改行を指定した場合、改行コードで分割された配列を返却。
'>NewLineCodeで改行を指定しなかった場合、文字列を返却。
'>
'>**Remarks**
'>
'>一括読み込み<br>
'>ファイルを一度に読み込みます。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Function IReader_ReadAll() As Variant

    IReader_ReadAll = mBuf

End Function
Private Sub Class_Terminate()

    On Error Resume Next

    '開いている場合にはクローズ
    If mOpen Then
        Close fp
    End If

End Sub
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
    If IsEmpty(mBuf) Then
        ICursor_Eof = True
    Else
        ICursor_Eof = UBound(mBuf) < mIndex
    End If
End Property
'------------------------------------------------------
' 現在の行取得
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    If ICursor_Eof Then
        ICursor_Item = mBuf(UBound(mBuf))
    Else
        ICursor_Item = mBuf(mIndex)
    End If
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    If IsEmpty(mBuf) Then
        mIndex = 0
    Else
        mIndex = LBound(mBuf)
    End If
End Sub
'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mIndex = mIndex + 1
End Sub

