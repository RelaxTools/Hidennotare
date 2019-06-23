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
' [HecatonCheir] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/HecatonCheir
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
'-----------------------------------------------------------------------------------------------------
' 2018-12-06 Ver.1.0.0 新規作成 テキストファイルクラス
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements IWriter

Private UTF8 As UTF8Encoding

Private fp As Integer
Private mOpen As Boolean

'プロパティ保持
Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants
Private mMode As OpenModeConstants

Private Function IConstructor_Instancing(Args As Collection) As Object

    If Args.Count = 0 Then
        Exit Function
    End If
    
    'Default
    mNewLineCode = NewLineCodeConstants.CRLF
    mEncode = EncodeConstants.CP932
    mMode = OpenModeConstants.Output
    
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
    
    Set IConstructor_Instancing = Me


    Dim strFolder As String
    
    'フォルダの存在チェック
    strFolder = FileIO.GetParentFolderName(mFileName)
    
    If Not FileIO.FolderExists(strFolder) Then
        Err.Raise 512 + 1, TypeName(Me) & ".OpenFile", "フォルダ(" & strFolder & ")が存在しません。"
    End If
    
    
    If mEncode = EncodeConstants.UTF8 Then
        Set UTF8 = New UTF8Encoding
    End If

    Call fileOpen

End Function

Private Sub IWriter_Append(ByVal Args As Variant)
    
    Dim strLine As String
    Dim bytBuf() As Byte

    '改行
    Select Case mNewLineCode
        Case NewLineCodeConstants.CRLF
            strLine = Args & vbCrLf
        Case NewLineCodeConstants.LF
            strLine = Args & vbLf
    End Select
        
        'エンコード
    Select Case mEncode
        Case EncodeConstants.CP932 'SJIS
            bytBuf = VBA.StrConv(strLine, vbFromUnicode)
        Case EncodeConstants.UTF8
            bytBuf = UTF8.GetBytes(strLine)
    End Select

    Put #fp, , bytBuf

End Sub

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

    Set UTF8 = Nothing

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

