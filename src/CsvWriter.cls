VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CsvWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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
' 依存モジュール
'   IWriter.cls
'   TextWriter.cls
'   StringBuilder.cls
'-----------------------------------------------------------------------------------------------------
' 2018-12-06 Ver.1.0.0 新規作成 CSV出力クラス
'-----------------------------------------------------------------------------------------------------

Option Explicit
Implements IConstructor
Implements IWriter


Private IR As IWriter

Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants
Private mMode As OpenModeConstants
Private mBom As Boolean

Private mDelimiter As String
Private mQuotation As Boolean

Private Function IConstructor_Instancing(Args As Collection) As Object
    
    If Args.Count = 0 Then
        Exit Function
    End If
    
    'Default
    mNewLineCode = NewLineCodeConstants.CRLF
    mEncode = EncodeConstants.CP932
    mMode = OpenModeConstants.Output
    mBom = False
    
    mDelimiter = ","
    mQuotation = False
    
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
        mMode = Args(5)
    End If
    
    If Args.Count >= 6 Then
        mDelimiter = Args(6)
    End If
    
    If Args.Count >= 7 Then
        mQuotation = Args(7)
    End If
    
    'TextWriterに委譲する。
    Set IR = Constructor(New TextWriter, mFileName, mNewLineCode, mEncode, mMode, mBom)
    
    Set IConstructor_Instancing = Me
    
End Function

Private Sub Class_Terminate()
    Set IR = Nothing
End Sub
Private Sub IWriter_Append(ByVal Args As Variant)
    
    Dim c As Variant
    Dim sb As StringBuilder
    Set sb = New StringBuilder

    If IsArray(Args) Then
        For Each c In Args
            sb.Append AddQuot(c, mQuotation)
        Next
        'TextWriterに委譲する。
        IR.Append sb.ToJoin(mDelimiter)
    Else
        'TextWriterに委譲する。
        IR.Append AddQuot(Args, mQuotation)
    End If

End Sub

Private Sub IWriter_Flush()
    'TextWriterに委譲する。
    IR.Flush
End Sub
'囲み文字設定
Private Function AddQuot(ByVal v As String, ByVal blnQuot As Boolean) As String

    If Trim$(v) = "" Or blnQuot = False Then
        AddQuot = v
    Else
        AddQuot = """" & v & """"
    End If

End Function
