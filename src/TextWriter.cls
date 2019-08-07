VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWriter"
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
'>### TextWriter クラス
'>
'>**Remarks**
'>
'>Textファイル書込みクラス<br>
'>
'>**Dependent**
'>
'>* IConstructor
'>* IWriter
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
Implements IWriter

Private fp As Integer
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
'>Set ret = Constructor(New TextWriter,
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
'>|ret|Object|新しいインスタンス(TextWriter)|
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
'>Set IW = Constructor(New TextWriter, strFile)
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
Private Function IConstructor_Instancing(Args As Collection) As Object

    If Args.Count = 0 Then
        Exit Function
    End If
    
    'Default
    mNewLineCode = NewLineCodeConstants.CRLF
    mEncode = EncodeConstants.CP932
    mMode = OpenModeConstants.Output
    mBom = False
    
    If Args.Count >= 1 Then
        mFileName = Args(1)
    End If
    
    If Args.Count >= 2 Then
        mNewLineCode = Args(2)
    End If
    
    If Args.Count >= 3 Then
        mEncode = Args(3)
    End If
    
    If Args.Count >= 4 Then
        mMode = Args(4)
    End If
    
    If Args.Count >= 5 Then
        mBom = Args(5)
    End If
        
    'IWriter Interface を返却する
    Set IConstructor_Instancing = Me.GetIWriter

    Dim strFolder As String
    
    'フォルダの存在チェック
    strFolder = FileIO.GetParentFolderName(mFileName)
    
    If Not FileIO.FolderExists(strFolder) Then
        Err.Raise 512 + 1, TypeName(Me) & ".IConstructor_Instancing", "not exists folder. (" & strFolder & ")"
    End If

    Call fileOpen

    mFirst = True

End Function
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
    If mBom And mFirst And mMode = OpenModeConstants.Output Then
        'エンコード
        Select Case mEncode
            Case EncodeConstants.UTF8
                
                ReDim bytBom(0 To 2)
            
                bytBom(0) = &HEF
                bytBom(1) = &HBB
                bytBom(2) = &HBF
            
                Put #fp, , bytBom
            
            Case EncodeConstants.UTF16LE
                
                ReDim bytBom(0 To 1)
            
                bytBom(0) = &HFF
                bytBom(1) = &HFE
        
                Put #fp, , bytBom
        
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
        Case EncodeConstants.CP932 'SJIS
            bytBuf = VBA.StrConv(strLine, vbFromUnicode)
        Case EncodeConstants.UTF8
            bytBuf = Convert.ToUTF8(strLine)
        Case EncodeConstants.UTF16LE
            bytBuf = strLine
    End Select

    Put #fp, , bytBuf

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
    If mBom And mFirst And mMode = OpenModeConstants.Output Then
        'エンコード
        Select Case mEncode
            Case EncodeConstants.UTF8
                
                ReDim bytBom(0 To 2)
            
                bytBom(0) = &HEF
                bytBom(1) = &HBB
                bytBom(2) = &HBF
            
                Put #fp, , bytBom
            
            Case EncodeConstants.UTF16LE
                
                ReDim bytBom(0 To 1)
            
                bytBom(0) = &HFF
                bytBom(1) = &HFE
        
                Put #fp, , bytBom
        
        End Select
        mFirst = False
    End If

    '改行
    Select Case mNewLineCode
        Case NewLineCodeConstants.CRLF
            strLine = Args & vbCrLf
        Case NewLineCodeConstants.LF
            strLine = Args & vbLf
        Case NewLineCodeConstants.None
            strLine = Args
    End Select
        
    'エンコード
    Select Case mEncode
        Case EncodeConstants.CP932 'SJIS
            bytBuf = VBA.StrConv(strLine, vbFromUnicode)
        Case EncodeConstants.UTF8
            bytBuf = Convert.ToUTF8(strLine)
        Case EncodeConstants.UTF16LE
            bytBuf = strLine
    End Select

    Put #fp, , bytBuf

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
        Close fp
    End If
    
    fp = FreeFile
    Open mFileName For Binary As fp
    mOpen = True

    Seek #fp, VBA.LOF(fp) + 1
    
End Sub

Private Sub Class_Terminate()

    On Error Resume Next

    '開いている場合にはクローズ
    If mOpen Then
        Close fp
    End If

End Sub

Private Sub fileOpen()
    
    Select Case mMode
        Case OpenModeConstants.Output
            
            fp = FreeFile()
            Open mFileName For Output As fp
            Close fp
        
            fp = FreeFile
            Open mFileName For Binary As fp
            mOpen = True
        
        Case OpenModeConstants.AppendFile
            
            fp = FreeFile
            Open mFileName For Binary As fp
            mOpen = True
        
            Seek #fp, VBA.LOF(fp) + 1
    
    End Select

End Sub
'>---
'>#### GetIWriter(TextWriter)
'>
'>**Syntax**
'>
'>```
'>Set i/f = obj.GetIWriter()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|i/f|IWriter|IWriter インターフェース(TextWriter)を返却|
'>
'>**Remarks**
'>
'>IWriterインターフェース返却メソッド<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function GetIWriter() As IWriter
    Set GetIWriter = Me
End Function

