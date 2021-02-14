VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvWriter"
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
'>### CsvWriter クラス
'>
'>**Remarks**
'>
'>CSV書込みクラス<br>
'>RFC4180準拠<br>
'>http://www.kasai.fm/wiki/rfc4180jp
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* StringBuilder
'>* IWriter
'>* ICsvWriterParam
'>* TextWriter
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
Implements IWriter
Implements ICsvWriterParam

Private mIW As IWriter

Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants
Private mMode As OpenModeConstants
Private mBom As Boolean

Private mDelimiter As String
Private mQuotation As Boolean
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = CsvWriter.NewInstance(FileName,
'>                                 [NewLineCode = NewLineCodeConstants.CRLF],
'>                                 [Encode = EncodeConstants.CP932],
'>                                 [Mode = OpenModeConstants.デフォルトOutput],
'>                                 [Bom = False],
'>                                 [Delimiter = ","])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|FileName|必須|String|ファイル名を指定|
'>|NewLineCode|省略可|String|改行コードを指定。<br>デフォルト(CRLF)|
'>|Encode|省略可|String|文字コードを指定。CP932/UTF8/UTF16LE<br>デフォルトCP932(SJIS)|
'>|Mode|省略可|OpenModeConstants|Output:上書きモード/Append:追加モード<br>デフォルトOutput|
'>|Bom|省略可|String|True:付ける/False:付けない<br>デフォルトFalse|
'>|Delimiter|省略可|String|区切り文字を指定。<br>デフォルト","|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|新しいインスタンス(IWriter)|
'>
'>**Remarks**
'>
'>** このクラスはコンストラクタを使用しないと実行できません。**
'>
'>**Example**
'>
'>
'>```
'>Dim strFile As String
'>Dim IW As IWriter
'>
'>strFile = "C:\Test.csv"
'>
'>Set IW = CsvWriter.NewInstance(strFile)
'>
'>IW.Append Array("a", "b", "c")
'>IW.Flash
'>
'>```
'>
'>**See also**
'>
'>* None
'>
'--------------------------------------------------------------
Public Function NewInstance(ByVal FileName As String, _
                             Optional ByVal NewLineCode As NewLineCodeConstants = NewLineCodeConstants.NewLineCodeCRLF, _
                             Optional ByVal Encode As EncodeConstants = EncodeConstants.EncodeCP932, _
                             Optional ByVal Mode As OpenModeConstants = OpenModeConstants.OpenModeOutput, _
                             Optional ByVal Bom As Boolean = False, _
                             Optional ByVal Delimiter As String = ",") As IWriter
    
    Dim FP As ICsvWriterParam
    Set FP = New CsvWriter
    
    FP.FileName = FileName
    FP.NewLineCode = NewLineCode
    FP.Encode = Encode
    FP.Mode = Mode
    FP.Bom = Bom
    FP.Delimiter = Delimiter
    
    FP.Init
    
    Set NewInstance = FP
    
End Function

Private Property Let ICsvWriterParam_Bom(RHS As Boolean)
    mBom = RHS
End Property

Private Property Let ICsvWriterParam_Delimiter(RHS As String)
    mDelimiter = RHS
End Property

Private Property Let ICsvWriterParam_Encode(RHS As EncodeConstants)
    mEncode = RHS

End Property

Private Property Let ICsvWriterParam_FileName(RHS As String)
    mFileName = RHS
End Property

Private Sub ICsvWriterParam_Init()
    'TextWriterに委譲する。
    Set mIW = TextWriter.NewInstance(mFileName, mNewLineCode, mEncode, mMode, mBom)
End Sub

Private Property Let ICsvWriterParam_Mode(RHS As OpenModeConstants)
    mMode = RHS
End Property

Private Property Let ICsvWriterParam_NewLineCode(RHS As NewLineCodeConstants)
    mNewLineCode = RHS

End Property

Private Sub Class_Terminate()
    Set mIW = Nothing
End Sub

Private Sub IWriter_WriteData(ByVal Args As Variant)
    Err.Raise 512 + 1, TypeName(Me) & ".IWriter_WriteData", "Not Support Method"
End Sub

'>---
'>#### WriteLine
'>
'>**Syntax**
'>
'>```
'>expression.WriteLine(Args)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|CsvWriterオブジェクトを指定します。|
'>|Args|必須|Variant|行イメージの配列を指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>行追加<br>
'>行イメージの配列をCSVとして出力します。<br>
'>明示的に保存する際にはFlushメソッドを呼び出すかインスタンスを終了(Set xxx = Nothing)する必要があります。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* CsvWriter.Flush
'>
Private Sub IWriter_WriteLine(ByVal Args As Variant)
    
    Dim c As Variant
    Dim SB As IStringBuilder
    Set SB = StringBuilder.NewInstance

    If IsArray(Args) Then
        For Each c In Args
            SB.Append AddQuot(c)
        Next
        'TextWriterに委譲する。
        mIW.WriteLine SB.ToString(mDelimiter)
    Else
        'TextWriterに委譲する。
        mIW.WriteLine AddQuot(Args)
    End If

End Sub
'>---
'>#### Flush
'>
'>**Syntax**
'>
'>```
'>expression.Flush(Args)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|StringBulder|CsvWriterオブジェクトを指定します。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>強制保存<br>
'>今までAppendで出力されたデータを強制的に保存します。<br>
'>インスタンスをモジュールレベルにするなど、長い間Appendしなしなかった場合に確実に保存される保証がありません。<br>
'>長い間Openしたままにする場合には、定期的にFlushしてください。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Private Sub IWriter_Flush()
    'TextWriterに委譲する。
    mIW.Flush
End Sub
'囲み文字設定
Private Function AddQuot(ByVal v As String) As String

    'デリミタまたはコーテーションまたは改行を含む場合
    Select Case True
        Case InStr(v, mDelimiter) > 0, InStr(v, vbLf) > 0
            AddQuot = """" & v & """"
        Case InStr(v, """") > 0
            AddQuot = """" & Replace(v, """", """""") & """"
        Case Else
            AddQuot = v
    End Select

End Function
