VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'>### CsvReader クラス
'>
'>**Remarks**
'>
'>CSV読込クラス<br>
'>RFC4180準拠<br>
'>http://www.kasai.fm/wiki/rfc4180jp<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* ICursor
'>* IReader
'>* TextReader
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements ICursor
Implements IReader

Private mIR As IReader

Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants

Private mDelimiter As String
Private mQuotation As Boolean

Private mCol As Collection
Private mIndex As Long

'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = CsvReader.CreateObject(
'>                    FileName,
'>                    [NewLineCode = NewLineCodeConstants.CRLF],
'>                    [Encode = EncodeConstants.CP932],
'>                    [Delimiter = ","],
'>                    [Quotation = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を指定|
'>|NewLineCode|省略可|String|改行コードを指定。<br>デフォルト(CRLF)|
'>|Encode|省略可|String|文字コードを指定。<br>デフォルトCP932(SJIS)|
'>|Delimiter|省略可|String|区切り文字を指定。<br>デフォルト","|
'>|Quotation|省略可|String|コーテーションの有無を指定。<br>デフォルトFalse|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(ICursor)|
'>
'>**Remarks**
'>
'>**このクラスはコンストラクタを使用しないと実行できません。**
'>
'>**Example**
'>
'>1行ずつ読む場合(ICursorインターフェイス)
'>
'>```
'>Dim strFile As String
'>Dim IC As ICursor
'>
'>strFile = "C:\Test.csv"
'>
'>Set IC = CsvReader.CreateObject(strFile)
'>
'>Do Until IC.Eof
'>
'>    Debug.Print IC.Item(1) '1列目
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
'>strFile = "C:\Test.csv"
'>
'>Set IR = CsvReader.CreateObject(strFile)
'>
'>Debug.Print IR.ReadAll '改行を指定した場合は配列、指定しなかった場合は文字列
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function CreateObject(ByVal FileName As String, _
                            Optional ByVal NewLineCode As NewLineCodeConstants = NewLineCodeConstants.NewLineCodeCRLF, _
                            Optional ByVal Encode As EncodeConstants = EncodeConstants.EncodeCP932, _
                            Optional ByVal Delimiter As String = ",", _
                            Optional ByVal Quotation As Boolean = False) As ICursor

    Dim CR As CSVReader
    
    Set CR = New CSVReader
    
    CR.FileName = FileName
    CR.NewLineCode = NewLineCode
    CR.Encode = Encode
    CR.Delimiter = Delimiter
    CR.Quotation = Quotation
    
    CR.Init

    Set CreateObject = CR

End Function
Public Property Let FileName(v As String)
    mFileName = v
End Property
Public Property Let NewLineCode(v As NewLineCodeConstants)
    mNewLineCode = v
End Property
Public Property Let Encode(v As EncodeConstants)
    mEncode = v
End Property
Public Property Let Delimiter(v As String)
    mDelimiter = v
End Property
Public Property Let Quotation(v As Boolean)
    mQuotation = v
End Property
Public Sub Init()
    
    'TextReaderに委譲する。改行コードをCsvReader側でパースする。
    Set mIR = TextReader.CreateObject(mFileName, NewLineCodeConstants.NewLineCodeNone, mEncode)
    
    If IsEmpty(mIR.ReadAll) Then
        Set mCol = New Collection
    Else
        Set mCol = Parser.ParseCsv(mIR.ReadAll(LBound(mIR.ReadAll)), mQuotation, mDelimiter)
    End If
    
    ICursor_MoveFirst

End Sub
Private Sub Class_Terminate()
    Set mIR = Nothing
End Sub
'>---
'>#### ReadAll(IReader I/F)
'>
'>**Syntax**
'>
'>```
'>set ret = expression.ReadAll
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|CsvReader|CsvReaderのインスタンスを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Collection|読み込んだデータを返却|
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

    Set IReader_ReadAll = mCol

End Function
'>---
'>#### Paste(CSvReader)
'>
'>**Syntax**
'>
'>```
'>expression.Paste(TopLeft, Fields , [Header = None])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|CsvReader|CsvReaderのインスタンスを指定|
'>|TopLeft|必須|Range|基準となる左上のセルを指定|
'>|Fields|必須|IDictionary|フィールド名を設定したIDictionary I/F に対応したクラスを指定|
'>|Header|省略可|CsvReaderHeaderConstants|AddHedder:ヘッダを追加/UseFirstRow:1行目をヘッダにする/None:ヘッダ無し<br>デフォルトNone|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>シート貼り付け<br>
'>読み込んだCSVをシートに貼り付けます。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub Paste(ByRef TopLeft As Range, ByRef Fields As IDictionary, Optional ByVal Header As CsvReaderHeaderConstants = CsvReaderHeaderConstants.HeaderNone)

    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim Key As Variant
    Const DEFAULT_FIELD As String = "Field"
    
    lngCol = 0
    lngRow = mCol.Count
    
    '列数がまちまちの場合があるので最大列数を調べる。
    For i = 1 To mCol.Count
        If lngCol < mCol(i).Count Then
            lngCol = mCol(i).Count
        End If
    Next
    
    '設定したフィールドの個数が一致しない場合エラー
    If lngCol <> Fields.Count Then
        Err.Raise 512 + 1, TypeName(Me) & ".Paste", "not match fields count. "
    End If
    
    Dim var() As Variant
    
    Select Case Header
        'ヘッダの追加
        Case CsvReaderHeaderConstants.AddHedder
            
            ReDim var(1 To lngRow + 1, 1 To lngCol)
            
            For j = 1 To lngCol
                If j <= Fields.Count Then
                    var(1, j) = Fields.Keys(j - 1)
                Else
                    var(1, j) = Default & Format$(j, "00")
                End If
            Next
            
            For i = 1 To mCol.Count
                For j = 1 To mCol(i).Count
                    var(i + 1, j) = mCol(i)(j)
                Next
            Next
        
        Case Else
            
            ReDim var(1 To lngRow, 1 To lngCol)
            
            For i = 1 To mCol.Count
                For j = 1 To mCol(i).Count
                    var(i, j) = mCol(i)(j)
                Next
            Next
    
    End Select
    
    Dim r As Range
    Set r = TopLeft.resize(lngRow, lngCol)
    
    Select Case Header
        '１行目をヘッダにする
        Case CsvReaderHeaderConstants.UseFirstRow
            
            For j = 1 To lngCol
                Key = var(j, 1)
                If Fields.ContainsKey(Key) Then
                    r.Columns(i).NumberFormatLocal = Fields(Key)
                End If
            Next
        
        'ヘッダを追加する
        Case CsvReaderHeaderConstants.AddHedder
            
            For j = 1 To lngCol
                If j <= Fields.Count Then
                    Key = Fields.Keys(j - 1)
                    If Fields.ContainsKey(Key) Then
                        r.Columns(j).NumberFormatLocal = Fields(Key)
                    End If
                End If
            Next
        
        'ヘッダ無し
        Case CsvReaderHeaderConstants.HeaderNone
            
            For j = 1 To lngCol
                Key = DEFAULT_FIELD & Format(j, "00")
                If Fields.ContainsKey(Key) Then
                    r.Columns(i).NumberFormatLocal = Fields(Key)
                End If
            Next
    End Select
    
    r.Value = var
    
End Sub
'>---
'>#### GetCursor(CsvReader)
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
'>|expression|必須|CsvReader|CsvReaderを指定します。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|var|ICursor|ICursor I/F を返却|
'>
'>**Remarks**
'>
'>CsvReaderのメソッドです。<br>
'>ICursor インターフェースを取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* ICursor
'>
Public Property Get GetCursor() As ICursor
    Set GetCursor = Me
End Property
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
    
    If mCol Is Nothing Then
        ICursor_Eof = True
        Exit Sub
    End If
    
    ICursor_Eof = mCol.Count < mIndex

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

    'EOF後は最後の値を返す
    If ICursor_Eof Then
        Core.SetVar(ICursor_Item) = mCol.Item(mCol.Count)(opt)
    Else
        Core.SetVar(ICursor_Item) = mCol.Item(mIndex)(opt)
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
    mIndex = 1
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


