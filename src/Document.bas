Attribute VB_Name = "Document"
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
'>### Document 標準モジュール
'>
'>**Remarks**
'>
'>* ドキュメント生成モジュール(Hidennotareをgitやwikiで管理するためのモジュール)
'>
'-----------------------------------------------------------------------------------------------------
Private Const TARGET_URL As String = "https://github.com/RelaxTools/Hidennotare/wiki/"
Option Explicit
'-----------------------------------------------------------------------------------------------------
' ソースのエクスポート
'-----------------------------------------------------------------------------------------------------
Sub Export()

    Dim strFile As String
    Dim strExt As String
    Dim obj As Object
    Dim strTo As String
    
    On Error GoTo e
    
    strFile = FileIO.BuildPath(FileIO.GetParentFolderName(ThisWorkbook.FullName), "src")
    FileIO.CreateFolder strFile
    
    For Each obj In ThisWorkbook.VBProject.VBComponents
    
        If obj.Name Like "Module*" Then
            GoTo pass
        End If
    
        Select Case obj.Type
            Case 1
                strExt = ".bas"
            Case 3
                strExt = ".frm"
            Case 2
                strExt = ".cls"
            Case 11, 100
                GoTo pass
        End Select
        
        strTo = FileIO.BuildPath(strFile, obj.Name & strExt)
        obj.Export strTo
pass:
    Next
    
    MsgBox "生成しました！", vbInformation, "Markdown"
    
    Exit Sub
e:
    If Err.Number = 70 Then
        If Message.Question("他のプログラムで開いています。再試行しますか？") Then
            Message.Critical "処理を中断しました。"
        Else
            Resume
        End If
    End If
    
End Sub
'-----------------------------------------------------------------------------------------------------
' Markdown出力
' Markdownがある行の先頭に「'>」があるものについてファイルに出力する。
'-----------------------------------------------------------------------------------------------------
Sub OutputMarkDown()
    
    Dim obj As Object
    Dim strFolder As String
    Dim strFile As String
    Dim SB As StringBuilder
    Dim strBuf As String
    Dim strMark As String
    Dim i As Long
    Dim TC As IList
    Dim fp As Integer
    Dim bytBuf() As Byte
    
    On Error GoTo e
    
    '目次作成用
    Set TC = New ArrayList
    
    '章番号を付加するレベル
    Const Level As Long = 4
    
    '目次を作成するレベル
    Const ContentsLevel As Long = 3
    
    Dim No1() As Long
    Dim No2() As Long
    Dim No3() As Long

    ReDim No1(1 To Level)
    ReDim No2(1 To Level)
    ReDim No3(1 To Level)

    '標準モジュールのスタート
     No1(1) = 2
     No1(2) = 1
     No1(3) = 0

    'インターフェイスのスタート
     No2(1) = 2
     No2(2) = 2
     No2(3) = 0

    'クラスのスタート
     No3(1) = 2
     No3(2) = 3
     No3(3) = 0
    
    'Hidennotare.wiki フォルダを作成
    strFolder = ThisWorkbook.Path & ".wiki"
    FileIO.CreateFolder strFolder
    
    
    'VBComponents の取得順が アルファベット順ではないので、SortedDictionary を使用。
    Dim dic As IDictionary
    Set dic = New SortedDictionary
    
    For Each obj In ThisWorkbook.VBProject.VBComponents
        dic.Add obj.Name, obj
    Next
    
    Dim Key As Variant
    For Each Key In dic.Keys
        
        Set obj = dic.Item(Key)
        
        'モジュール名.md を作成する。
        strFile = FileIO.BuildPath(strFolder, obj.Name & ".md")
        
        With obj.CodeModule
            
            Set SB = New StringBuilder
            
            For i = 1 To .CountOfLines
                
                '指定位置から１行取得
                strBuf = .Lines(i, 1)
                
                If Left$(strBuf, 2) = "'>" Then
                    
                    '------------------------------------------
                    ' 章番号の生成
                    '------------------------------------------
                    strMark = Mid$(strBuf, 3)
                    Select Case True
                        
                        '標準モジュール
                        Case obj.Type = 1
                           SB.Append LevelNo(strMark, No1(), Level, TC, ContentsLevel, obj.Name)
                        
                        '1文字目が"I"、2文字目が大文字の場合、インターフェース
                        Case RegExp.Test(obj.Name, "^I[A-Z]")
                           SB.Append LevelNo(strMark, No2(), Level, TC, ContentsLevel, obj.Name)
                        
                        'その他クラス
                        Case Else
                           SB.Append LevelNo(strMark, No3(), Level, TC, ContentsLevel, obj.Name)
                    
                    End Select
                End If
            Next i
        
            '対象があれば出力する
            If SB.Length > 0 Then
                
                'ファイルを空にする。
                FileIO.TruncateFile strFile
                
                'UTF8 & LF で保存
                fp = FreeFile()
                Open strFile For Binary As fp
                
                bytBuf = Convert.ToUTF8(SB.ToJoin(vbLf))
                
                Put #fp, , bytBuf
                Close fp
            End If
            
            Set SB = Nothing
        
        End With
    
    Next
    
    'Wikiの目次作成
    If TC.Count > 0 Then
    
        Dim strStatic As String
        
        'Wikiの目次ファイル名
        strFile = FileIO.BuildPath(strFolder, "_Sidebar.md")
        
        '------------------------------------------
        ' 目次の静的コンテンツ部分を取得
        '------------------------------------------
        strStatic = GetStaticContents(strFile)
        
        'ソート
        TC.Sort New ExplorerComparer
        
        '目次作成
        TC.Insert 0, "#### 2 リファレンス"
        TC.Insert 1, "##### 2.1 標準モジュール"
        For i = 0 To TC.Count
            If StringHelper.StartsWith(TC.Item(i), "[2.2") Then
                TC.Insert i, "##### 2.2 インターフェイス"
                Exit For
            End If
        Next
        For i = 0 To TC.Count
            If StringHelper.StartsWith(TC.Item(i), "[2.3") Then
                TC.Insert i, "##### 2.3 クラス"
                Exit For
            End If
        Next
        
        '元ファイルをクリア
        FileIO.TruncateFile strFile
        
        '静的コンテンツと生成した目次をUTF8 & LF にて出力。
        fp = FreeFile()
        Open strFile For Binary As fp
        
        bytBuf = Convert.ToUTF8(strStatic)
        
        Put #fp, , bytBuf
        
        bytBuf = Convert.ToUTF8(Join(TC.ToArray(), vbLf))
        
        Put #fp, , bytBuf
        Close fp
    
    End If
    
    MsgBox "生成しました！", vbInformation, "Markdown"
    
    Exit Sub
e:
    If Err.Number = 70 Then
        If Message.Question("他のプログラムで開いています。再試行しますか？") Then
            Message.Critical "処理を中断しました。"
        Else
            Resume
        End If
    End If
End Sub
'---------------------------------------------------
' 章番号生成
'---------------------------------------------------
Private Function LevelNo(ByVal strBuf As String, No() As Long, ByVal lngLevel As Long, TC As IList, ByVal lngContentsLevel As Long, ByVal strName As String) As String

    Dim Col As Collection
    Dim SB As StringBuilder
    Dim lngLen As Long
    Dim i As Long
    
    '章番号(###～)の場合
    Set Col = RegExp.Execute(strBuf, "^#+ ")

    If Col.Count > 0 Then
    
        lngLen = Len(Col(1).Value) - 1
        
        Dim strLeft As String
        Dim strRight As String
        
        strLeft = Col(1).Value
        strRight = Mid$(strBuf, Col(1).Length)
        
        '章番号生成レベル以上であれば、章番号作成
        If lngLen <= lngLevel Then

            '章番号をカウントアップ
            No(lngLen) = No(lngLen) + 1

            '現レベル以下の番号をクリア
            For i = lngLen + 1 To lngLevel
                No(i) = 0
            Next

            '章番号の生成
            Set SB = New StringBuilder
            For i = 1 To lngLen
                SB.Append CStr(No(i))
            Next
            
            LevelNo = strLeft & SB.ToJoin(".") & strRight
        
        
            '目次作成レベル以上であれば目次作成
            If lngLen <= lngContentsLevel Then
            
                Dim strContent As String
                
                'Markdown のリンク見出しとリンク作成
                strContent = "[" & Mid$(LevelNo, InStr(LevelNo, " ") + 1) & "](" & TARGET_URL & Replace$(strName, " ", "-") & ")  "
                
                '目次のエリアが限られるので種類は削除
                strContent = Replace$(strContent, " クラス", "")
                strContent = Replace$(strContent, " インターフェイス", "")
                strContent = Replace$(strContent, " 標準モジュール", "")
            
                TC.Add strContent
            
            End If
        
        Else
            LevelNo = strBuf
        End If
    Else
        LevelNo = strBuf
    End If

End Function
'---------------------------------------------------
'目次から静的コンテンツ部分を抜き出す
'---------------------------------------------------
Private Function GetStaticContents(ByVal strFile As String) As String

    Dim fp As Integer
    Dim bytBuf() As Byte
    Dim strBuf As String
    
    GetStaticContents = ""
    
    fp = FreeFile
    
    Open strFile For Binary As fp
    
    ReDim bytBuf(0 To LOF(fp) - 1)
    
    Get #fp, , bytBuf

    Close fp
    
    strBuf = Convert.FromUTF8(bytBuf)

    Dim SB As StringBuilder
    
    Set SB = New StringBuilder

    Dim IC As ICursor
    
    Set IC = Constructor(New LineCursor, Split(strBuf, vbLf))
    
    Do Until IC.Eof
    
        If StringHelper.StartsWith(IC, "#### 2") Then
            Exit Do
        End If
        
        SB.Append IC
        
        IC.MoveNext
    Loop

    If SB.Length > 0 Then
        GetStaticContents = SB.ToJoin(vbLf) & vbLf
    End If
    
End Function
