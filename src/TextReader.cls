VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextReader"
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
'>### TextReader クラス
'>
'>**Remarks**
'>
'>TextReaderクラス<br>
'>一度にメモリに読み込むため、巨大なファイルを読むには向かない。ファイルサイズ10MB以上非推奨。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>- IReader
'>- ITextReaderParam
'>- ICursor
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにNewInstance以外表示されないように修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IReader
Implements ICursor
Implements IParam

Private FP As Integer
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
'>Set IC = TextReader.NewInstance(FileName,
'>                                 [NewLineCode = NewLineCodeConstants.CRLF],
'>                                 [Encode = EncodeConstants.CP932])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を設定|
'>|NewLineCode|省略可|NewLineCodeConstants|読み取り専用の設定。<br>デフォルトCRLF|
'>|Encode|省略可|EncodeConstants|Applicationオブジェクト。<br>デフォルトCP932|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|IC|Object|新しいインスタンス(ICursor)|
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
'>strFile = "C:\Test.txt"
'>
'>Set IC = TextReader.NewInstance(strFile)
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
'>strFile = "C:\Test.txt"
'>
'>Set IR = TextReader.NewInstance(strFile)
'>
'>Debug.Print IR.ReadAll '改行を指定した場合は配列、指定しなかった場合は文字列
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function NewInstance(ByVal FileName As String, _
                            Optional ByVal NewLineCode As NewLineCodeConstants = NewLineCodeConstants.NewLineCodeCRLF, _
                            Optional ByVal Encode As EncodeConstants = EncodeConstants.EncodeCP932) As ICursor

    Dim p As IParam
    
    Set p = New TextReader
    
    p.Parameters = Array(FileName, NewLineCode, Encode)
    
    Set NewInstance = p

End Function
Private Property Let IParam_Parameters(ByVal Args As Variant)

    mFileName = Args(0)
    mNewLineCode = Args(1)
    mEncode = Args(2)
    
    BulkRead
    
    ICursor_MoveFirst

End Property

'--------------------------------------------------------------
' テキストファイルの一括読み込み
'--------------------------------------------------------------
Private Sub BulkRead()
    
    Dim Buffer() As Byte
    Dim strBuf As String
    Dim lngSize As Long

    If Not FileIO.FileExists(mFileName) Then
        Err.Raise 512 + 1, TypeName(Me) & ".BulkRead", "file not found(" & mFileName & ")"
    End If

    FP = FreeFile
    Open mFileName For Binary As FP
    mOpen = True
    
    lngSize = VBA.LOF(FP)
    
    If lngSize <> 0 Then

        ReDim Buffer(0 To lngSize - 1)
    
        Get FP, , Buffer()
        
        Select Case mEncode
            
            Case EncodeConstants.EncodeCP932
                
                strBuf = StrConv(Buffer, vbUnicode)
            
            Case EncodeConstants.EncodeUTF8
            
                'BOMがある場合取り除く
                If StartWith(Buffer, &HEF, &HBB, &HBF) Then
                    strBuf = StringUtils.FromUTF8(MidB$(Buffer, 4))
                Else
                    strBuf = StringUtils.FromUTF8(Buffer)
                End If
        
            Case EncodeConstants.EncodeUTF16LE
                
                'BOMがある場合取り除く
                If StartWith(Buffer, &HFF, &HFE) Then
                    strBuf = MidB$(Buffer, 3)
                Else
                    strBuf = Buffer
                End If
        
        End Select
        
        Select Case mNewLineCode
            Case NewLineCodeConstants.NewLineCodeCRLF
                
                '最後に改行がある場合削除する。
                If Mid$(strBuf, Len(strBuf) - 1) = vbCrLf Then
                    strBuf = Mid$(strBuf, 1, Len(strBuf) - 2)
                End If
                
                mBuf = Split(strBuf, vbCrLf)
            
            Case NewLineCodeConstants.NewLineCodeLF
                
                '最後に改行がある場合削除する。
                If Mid$(strBuf, Len(strBuf)) = vbLf Then
                    strBuf = Mid$(strBuf, 1, Len(strBuf) - 1)
                End If
                
                mBuf = Split(strBuf, vbLf)
        
            Case NewLineCodeConstants.NewLineCodeNone
                
                mBuf = Array(strBuf)
        
        End Select

    Else
        mBuf = Empty
    End If
    
    Close FP
    mOpen = False
        
End Sub
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
'>* None
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
        Close FP
    End If

End Sub
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
    If IsEmpty(mBuf) Then
        ICursor_Eof = True
    Else
        ICursor_Eof = UBound(mBuf) < mIndex
    End If
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
        ICursor_Item = mBuf(UBound(mBuf))
    Else
        ICursor_Item = mBuf(mIndex)
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
    If IsEmpty(mBuf) Then
        mIndex = 0
    Else
        mIndex = LBound(mBuf)
    End If
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
End Sub

