VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVReader"
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
'   IConstructor.cls
'   ICursor.cls
'   IReader.cls
'   TextReader.cls
'-----------------------------------------------------------------------------------------------------
' 2018-12-06 Ver.1.0.0 新規作成 CSV入力クラス RFC4180準拠
' http://www.kasai.fm/wiki/rfc4180jp
'-----------------------------------------------------------------------------------------------------
' Dim IR As IReader
' Set IR = Constructor(new CsvReader, ファイル名, 改行コード, エンコード, デリミタ, コーテーション有無
'
' IR.ReadAll
'-----------------------------------------------------------------------------------------------------
' Dim IC As ICursor
' Set IC = Constructor(new CsvReader, ファイル名, 改行コード, エンコード, デリミタ, コーテーション有無
'
' Do Until IC.Eof
'     Debug.Print IC(1) '列
'     IC.MoveNext
' Loop
'-----------------------------------------------------------------------------------------------------
' Dim CR As CsvReader
' Set CR = Constructor(new CsvReader, ファイル名, 改行コード, エンコード, デリミタ, コーテーション有無
'
' CR.Paste 張り付けセルの左上, フィールド情報, ヘッダ設定
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor
Implements IReader

Private IR As IReader

Private mFileName As String
Private mNewLineCode As NewLineCodeConstants
Private mEncode As EncodeConstants

Private mDelimiter As String
Private mQuotation As Boolean

Private mList As Collection
Private mIndex As Long
Public Enum CsvReaderHeaderConstants
    UseFirstRow
    AddHedder
    None
End Enum
'--------------------------------------------------------------
'　コンストラクタ
'--------------------------------------------------------------
Private Function IConstructor_Instancing(Args As Collection) As Object
    
    If Args.Count = 0 Then
        Exit Function
    End If
    
    'Default
    mNewLineCode = NewLineCodeConstants.CRLF
    mEncode = EncodeConstants.CP932
    
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
        mDelimiter = Args(4)
    End If
    
    If Args.Count >= 5 Then
        mQuotation = Args(5)
    End If
    
    'TextReaderに委譲する。改行コードをCsvReader側でパースする。
    Set IR = Constructor(New TextReader, mFileName, NewLineCodeConstants.None, mEncode)
    
    Set mList = CsvParser(IR.ReadAll)
    
    ICursor_MoveFirst
    
    Set IConstructor_Instancing = Me
    
End Function
Private Sub Class_Terminate()
    Set IR = Nothing
End Sub
'--------------------------------------------------------------
'　CSVパーサー
'--------------------------------------------------------------
Private Function CsvParser(ByVal strBuf As String) As Collection

    Const C_QUAT As String = """"
    Dim IC As ICursor
    Dim sw As Boolean
    Dim blnLineBreak As Boolean
    Dim blnItemBreak As Boolean
    Dim col As Collection
    Dim row As Collection
    Dim lngQuot As Long
    
    lngQuot = 0
    
    sw = False
    
    Set row = New Collection
    
    Set IC = Constructor(New CharCursor, strBuf)
    Do Until IC.Eof
        
        '初期化
        Set col = New Collection
        blnLineBreak = False

        Do Until IC.Eof Or blnLineBreak
        
            '初期化
            Dim sb As StringBuilder
            Set sb = New StringBuilder
            blnItemBreak = False
            
            Do Until IC.Eof Or blnLineBreak Or blnItemBreak

                Select Case IC.Item
                    Case C_QUAT

                        'コーテーションありの場合
                        If mQuotation Then

                            lngQuot = lngQuot + 1

                            If sw Then

                                '次の文字がQuatation
                                If IC.Item(1) = C_QUAT Then
                                    lngQuot = lngQuot + 1
                                    sb.Append C_QUAT
                                    IC.MoveNext
                                Else
                                    sw = False
                                End If

                            Else
                                sw = True
                            End If
                        End If

                    Case vbCr

                        If sw Then
                            sb.Append IC.Item
                        End If

                    Case vbLf

                        If sw Then
                            sb.Append IC.Item
                        Else
                            blnLineBreak = True
                        End If

                    Case mDelimiter

                        If sw Then
                            sb.Append IC.Item
                        Else
                            blnItemBreak = True
                        End If

                    Case Else

                        sb.Append IC.Item

                End Select
            
                IC.MoveNext
            Loop
            
            '列追加処理
            col.Add sb.ToString
            
        Loop
        
        '行追加処理
        row.Add col
        
    Loop
    
    'ダブルコーテーションが偶数ではない場合
    If lngQuot Mod 2 <> 0 Then
        Message.Throw 1, Me, "CsvParser", "Invalid Format(Quotation)"
    End If
    
    Set CsvParser = row
    
End Function
'--------------------------------------------------------------
'　配列読み込み
'--------------------------------------------------------------
Private Function IReader_ReadAll() As Variant

    Set IReader_ReadAll = mList

End Function
'--------------------------------------------------------------
'　シート貼り付け
'--------------------------------------------------------------
Public Sub Paste(ByRef TopLeft As Range, ByRef Fields As Dictionary, Optional ByVal Header As CsvReaderHeaderConstants = CsvReaderHeaderConstants.None)

    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim Key As Variant
    Const DEFAULT_FIELD As String = "Field"
    
    lngCol = 0
    lngRow = mList.Count
    
    '列数がまちまちの場合があるので最大列数を調べる。
    For i = 1 To mList.Count
        If lngCol < mList(i).Count Then
            lngCol = mList(i).Count
        End If
    Next
    
    '設定したフィールドの個数が一致しない場合エラー
    If lngCol <> Fields.Count Then
        Message.Throw 1, Me, "Paste", "fields count does not match"
    End If
    
    Dim var() As Variant
    
    Select Case Header
        'ヘッダの追加
        Case CsvReaderHeaderConstants.AddHedder
            
            ReDim var(1 To lngRow + 1, 1 To lngCol)
            
            For j = 1 To lngCol
                If j <= Fields.Count Then
                    var(1, j) = Fields.keys(j - 1)
                Else
                    var(1, j) = Default & Format$(j, "00")
                End If
            Next
            
            For i = 1 To mList.Count
                For j = 1 To mList(i).Count
                    var(i + 1, j) = mList(i)(j)
                Next
            Next
        
        Case Else
            
            ReDim var(1 To lngRow, 1 To lngCol)
            
            For i = 1 To mList.Count
                For j = 1 To mList(i).Count
                    var(i, j) = mList(i)(j)
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
                If Fields.Exists(Key) Then
                    r.Columns(i).NumberFormatLocal = Fields(Key)
                End If
            Next
        
        'ヘッダを追加する
        Case CsvReaderHeaderConstants.AddHedder
            
            For j = 1 To lngCol
                If j <= Fields.Count Then
                    Key = Fields.keys(j - 1)
                    If Fields.Exists(Key) Then
                        r.Columns(j).NumberFormatLocal = Fields(Key)
                    End If
                End If
            Next
        
        'ヘッダ無し
        Case CsvReaderHeaderConstants.None
            
            For j = 1 To lngCol
                Key = DEFAULT_FIELD & Format(j, "00")
                If Fields.Exists(Key) Then
                    r.Columns(i).NumberFormatLocal = Fields(Key)
                End If
            Next
    End Select
    
    r.Value = var
    
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
    
    If mList Is Nothing Then
        ICursor_Eof = True
        Exit Sub
    End If
    
    ICursor_Eof = mList.Count < mIndex

End Property
'------------------------------------------------------
' 項目取得
'------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant

    If ICursor_Eof Then
        'EOF後は最後の値を返す
        If IsObject(mList.Item(mList.Count)(opt)) Then
            Set ICursor_Item = mList.Item(mList.Count)(opt)
        Else
            ICursor_Item = mList.Item(mList.Count)(opt)
        End If
    Else
        If IsObject(mList.Item(mIndex)(opt)) Then
            Set ICursor_Item = mList.Item(mIndex)(opt)
        Else
            ICursor_Item = mList.Item(mIndex)(opt)
        End If
    End If

End Property
'------------------------------------------------------
' 前項目取得
'------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant

    Message.Throw 1, Me, "ICursor_PreviousItem", "Not Supported Property"
    
End Property
'------------------------------------------------------
' 最初の行に移動
'------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = 1
End Sub

'------------------------------------------------------
' 次行取得
'------------------------------------------------------
Private Sub ICursor_MoveNext()
    mIndex = mIndex + 1
End Sub


