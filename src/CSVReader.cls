VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVReader"
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
'>* ICsvReaderParam
'>* TextReader
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
Implements ICursor
Implements IReader
Implements IParam

Private mIR As IReader

Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants

Private mDelimiter As String
Private mQuotation As Boolean
Private mHeader As Boolean
Private mSkip As Long

Private MList As IList
Private mIndex As Long

'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = CsvReader.NewInstance(
'>                    FileName,
'>                    [NewLineCode = NewLineCodeConstants.CRLF],
'>                    [Encode = EncodeConstants.CP932],
'>                    [Delimiter = ","],
'>                    [Quotation = False],)
'>                    [Header = False],)
'>                    [Skip = 0])
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
'>|Quotation|省略可|Boolean|コーテーションの有無を指定。<br>デフォルトFalse|
'>|Header|省略可|Boolean|ヘッダの有無を指定。指定すると列名が使える<br>デフォルトFalse|
'>|Skip|省略可|long|先頭から指定行分行をスキップする<br>デフォルト0|
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
'>Set IC = CsvReader.NewInstance(strFile)
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
'>Set IR = CsvReader.NewInstance(strFile)
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
                            Optional ByVal Encode As EncodeConstants = EncodeConstants.EncodeCP932, _
                            Optional ByVal Delimiter As String = ",", _
                            Optional ByVal Quotation As Boolean = False, _
                            Optional ByVal Header As Boolean = False, _
                            Optional ByVal Skip As Long = 0) As ICursor

    Dim p As IParam
    
    Set p = New CSVReader
    
    p.Parameters = Array(FileName, NewLineCode, Encode, Delimiter, Quotation, Header, Skip)
    
    Set NewInstance = p

End Function
Private Property Let IParam_Parameters(ByVal Args As Variant)
    
    mFileName = Args(0)
    mNewLineCode = Args(1)
    mEncode = Args(2)
    mDelimiter = Args(3)
    mQuotation = Args(4)
    mHeader = Args(5)
    mSkip = Args(6)
    
    'TextReaderに委譲する。改行コードをCsvReader側でパースする。
    Set mIR = TextReader.NewInstance(mFileName, NewLineCodeConstants.NewLineCodeNone, mEncode)
    
    If IsEmpty(mIR.ReadAll) Then
        Set MList = ArrayList.NewInstance()
    Else
        Set MList = Parser.ParseCsv(mIR.ReadAll(LBound(mIR.ReadAll)), mQuotation, mDelimiter, mHeader, mSkip)
    End If
    
    ICursor_MoveFirst

End Property
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

    Set IReader_ReadAll = MList

End Function
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
    
    If MList Is Nothing Then
        ICursor_Eof = True
        Exit Sub
    End If
    
    ICursor_Eof = MList.Count <= mIndex

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

    
    'ヘッダー無の場合、IList/IList
    If Not mHeader Then
    
        Dim col As IList
        
        'EOF後は最後の値を返す
        If ICursor_Eof Then
            Set col = MList.Item(MList.Count - 1)
        Else
            Set col = MList.Item(mIndex)
        End If
    
        Objects.SetWith(ICursor_Item) = col.Item(CLng(opt))
    
    'ヘッダー有の場合、IList/IDictionary
    Else
        Dim pos As Long
        
        If ICursor_Eof Then
            pos = MList.Count - 1
        Else
            pos = mIndex
        End If
        
        Dim dic As IDictionary


        Set dic = MList.Item(pos)
    
        Select Case VarType(opt)
            '文字列
            Case vbString
                If dic.ContainsKey(opt) Then
                    Objects.SetWith(ICursor_Item) = dic(opt)
                Else
                    Err.Raise 512 + 1, TypeName(Me) & ".CSVReader", StringUtils.PlaceHolder("Invalid Name({0})", opt)
                End If
            '整数
            Case vbInteger, vbLong
                Select Case opt
                    Case LBound(dic.Keys) To UBound(dic.Keys)
                    Case Else
                        Err.Raise 512 + 1, TypeName(Me) & ".CSVReader", StringUtils.PlaceHolder("Invalid Index({0})", opt)
                End Select
            '上記以外
            Case Else
                Err.Raise 512 + 1, TypeName(Me) & ".CSVReader", "Invalid Type"
        End Select
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
    mIndex = 0
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


