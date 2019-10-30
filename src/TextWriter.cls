VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
'>### TextWriter クラス
'>
'>**Remarks**
'>
'>Textファイル書込みクラス<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* IWriter
'>* ITextWriterParam
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2018-12-06|Ver.1.0.0|新規作成|
'>|2019-10-28|Ver.1.0.1|クラスのメンバにCreateObject以外表示されないように修正|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IWriter
Implements ITextWriterParam

Private FP As Integer
Private mOpen As Boolean

'プロパティ保持
Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants
Private mMode As OpenModeConstants
Private mBom As Boolean

Private mFirst As Boolean
'>---
'>#### コンストラクタ
'>
'>**Syntax**
'>
'>```
'>Set ret = TextWriter.CreateObject(
'>                      FileName,
'>                      [NewLineCode = NewLineCodeConstants.CRLF],
'>                      [Encode = EncodeConstants.CP932],
'>                      [Mode = OpenModeConstants.デフォルトOutput],
'>                      [Bom = False])
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
'>strFile = "C:\Test.txt"
'>
'>Set IW = TextWriter.CreateObject(strFile)
'>
'>IW.Append "Hello World!"
'>IW.Flash
'>
'>```
'>
'>**See also**
'>
'>* None
'>
Public Function CreateObject(ByVal FileName As String, _
                             Optional ByVal NewLineCode As NewLineCodeConstants = NewLineCodeCRLF, _
                             Optional ByVal Encode As EncodeConstants = EncodeCP932, _
                             Optional ByVal Mode As OpenModeConstants = OpenModeOutput, _
                             Optional ByVal Bom As Boolean = False) As IWriter
    
    Dim FP As ITextWriterParam
    Set FP = New TextWriter
    
    FP.FileName = FileName
    FP.NewLineCode = NewLineCode
    FP.Encode = Encode
    FP.Mode = Mode
    FP.Bom = Bom
    
    FP.Init
    
    Set CreateObject = FP
    
End Function

Private Property Let ITextWriterParam_Bom(RHS As Boolean)
    mBom = RHS
End Property

Private Property Let ITextWriterParam_Encode(RHS As EncodeConstants)
    mEncode = RHS
End Property

Private Property Let ITextWriterParam_FileName(RHS As String)
    mFileName = RHS
End Property

Private Property Let ITextWriterParam_Mode(RHS As OpenModeConstants)
    mMode = RHS
End Property

Private Property Let ITextWriterParam_NewLineCode(RHS As NewLineCodeConstants)
    mNewLineCode = RHS
End Property

Private Sub ITextWriterParam_Init()
    fileOpen
End Sub
'>---
'>#### WriteData(IWriter I/F)
'>
'>**Syntax**
'>
'>```
'>expression.WriteData(Args)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|expression|必須|TextWriter|TextWriterオブジェクトを指定します。|
'>|Args|必須|Variant|テキストを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テキストの書込み（改行なし）<br>
'>指定データをテキストとして出力します。<br>
'>明示的に保存する際にはFlushメソッドを呼び出すかインスタンスを終了(Set xxx = Nothing)する必要があります。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* TextWriter.Flush
'>
Private Sub IWriter_WriteData(ByVal Args As Variant)
    
    Dim strLine As String
    Dim bytBuf() As Byte
    Dim bytBom() As Byte
    
    '初回BOMの書き込みを行う。
    If mBom And mFirst And mMode = OpenModeConstants.OpenModeOutput Then
        'エンコード
        Select Case mEncode
            Case EncodeConstants.EncodeUTF8
                
                ReDim bytBom(0 To 2)
            
                bytBom(0) = &HEF
                bytBom(1) = &HBB
                bytBom(2) = &HBF
            
                Put #FP, , bytBom
            
            Case EncodeConstants.EncodeUTF16LE
                
                ReDim bytBom(0 To 1)
            
                bytBom(0) = &HFF
                bytBom(1) = &HFE
        
                Put #FP, , bytBom
        
        End Select
        mFirst = False
    End If

'    '改行
'    Select Case mNewLineCode
'        Case NewLineCodeConstants.CRLF
'            strLine = Args & vbCrLf
'        Case NewLineCodeConstants.LF
'            strLine = Args & vbLf
'        Case NewLineCodeConstants.None
            strLine = Args
'    End Select
        
    'エンコード
    Select Case mEncode
        Case EncodeConstants.EncodeCP932  'SJIS
            bytBuf = VBA.StrConv(strLine, vbFromUnicode)
        Case EncodeConstants.EncodeUTF8
            bytBuf = Convert.ToUTF8(strLine)
        Case EncodeConstants.EncodeUTF16LE
            bytBuf = strLine
    End Select

    Put #FP, , bytBuf

End Sub
'>---
'>#### WriteLine(IWriter I/F)
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
'>|expression|必須|TextWriter|TextWriterオブジェクトを指定します。|
'>|Args|必須|Variant|テキストを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テキストの書込み（改行あり）<br>
'>指定データをテキストとして出力します（改行あり）。<br>
'>明示的に保存する際にはFlushメソッドを呼び出すかインスタンスを終了(Set xxx = Nothing)する必要があります。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* TextWriter.Flush
'>
Private Sub IWriter_WriteLine(ByVal Args As Variant)
    
    Dim strLine As String
    Dim bytBuf() As Byte
    Dim bytBom() As Byte
    
    '初回BOMの書き込みを行う。
    If mBom And mFirst And mMode = OpenModeConstants.OpenModeOutput Then
        'エンコード
        Select Case mEncode
            Case EncodeConstants.EncodeUTF8
                
                ReDim bytBom(0 To 2)
            
                bytBom(0) = &HEF
                bytBom(1) = &HBB
                bytBom(2) = &HBF
            
                Put #FP, , bytBom
            
            Case EncodeConstants.EncodeUTF8
                
                ReDim bytBom(0 To 1)
            
                bytBom(0) = &HFF
                bytBom(1) = &HFE
        
                Put #FP, , bytBom
        
        End Select
        mFirst = False
    End If

    '改行
    Select Case mNewLineCode
        Case NewLineCodeConstants.NewLineCodeCRLF
            strLine = Args & vbCrLf
        Case NewLineCodeConstants.NewLineCodeLF
            strLine = Args & vbLf
        Case NewLineCodeConstants.NewLineCodeNone
            strLine = Args
    End Select
        
    'エンコード
    Select Case mEncode
        Case EncodeConstants.EncodeCP932  'SJIS
            bytBuf = VBA.StrConv(strLine, vbFromUnicode)
        Case EncodeConstants.EncodeUTF8
            bytBuf = Convert.ToUTF8(strLine)
        Case EncodeConstants.EncodeUTF16LE
            bytBuf = strLine
    End Select

    Put #FP, , bytBuf

End Sub

'>---
'>#### Flush(IWriter I/F)
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
'>|expression|必須|TextWriter|TextWriterオブジェクトを指定します。|
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
    
    If mOpen = False Then Exit Sub

    On Error Resume Next
    
    If mOpen Then
        Close FP
    End If
    
    FP = FreeFile
    Open mFileName For Binary As FP
    mOpen = True

    Seek #FP, VBA.LOF(FP) + 1
    
End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    '開いている場合にはクローズ
    FileClose

End Sub
Private Sub fileOpen()
    
    Dim strFolder As String

    'フォルダの存在チェック
    strFolder = FileIO.GetParentFolderName(mFileName)

    If Not FileIO.FolderExists(strFolder) Then
        Err.Raise 512 + 1, TypeName(Me) & ".fileOpen", "not exists folder. (" & strFolder & ")"
    End If
    
    Select Case mMode
        Case OpenModeConstants.OpenModeOutput
            
            FP = FreeFile()
            Open mFileName For Output As FP
            Close FP
        
            FP = FreeFile
            Open mFileName For Binary As FP
            mOpen = True
        
        Case OpenModeConstants.OpenModeAppend
            
            FP = FreeFile
            Open mFileName For Binary As FP
            mOpen = True
        
            Seek #FP, VBA.LOF(FP) + 1
    
    End Select
    mFirst = True
End Sub
Private Sub FileClose()
    If mOpen Then
        Close FP
    End If
End Sub

