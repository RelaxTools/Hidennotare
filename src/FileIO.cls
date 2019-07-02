VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。
'-----------------------------------------------------------------------------------------------------
' 依存モジュール
'   Lang.bas
'   Registry.bas
'-----------------------------------------------------------------------------------------------------
' 2018-12-16 Ver.1.0.0 新規作成 CreateObjectしなくても使えるFileSystemObject
'-----------------------------------------------------------------------------------------------------
Option Explicit

Private Declare PtrSafe Function FindFirstFileExW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal fInfoLevelId As Long, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal fSearchOp As Long, ByVal lpSearchFilter As LongPtr, ByVal dwAdditionalFlags As Long) As LongPtr
Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As LongPtr, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare PtrSafe Function WNetGetConnectionW Lib "mpr" (ByVal lpszLocalName As LongPtr, ByVal lpszRemoteName As LongPtr, lSize As Long) As Long
Private Declare PtrSafe Function SearchTreeForFileW Lib "dbghelp" (ByVal RootPath As LongPtr, ByVal InputPathName As LongPtr, ByVal OutputPathBuffer As LongPtr) As Long
Private Declare PtrSafe Function StrCmpLogicalW Lib "Shlwapi" (ByVal psz1 As LongPtr, ByVal psz2 As LongPtr) As Long

Private Const MAX_PATH                  As Long = 260 * 2 - 1   'パスの最大長
Private Const MAX_PATH_SHORT            As Long = 14 * 2 - 1    'パスの最大長

Private Const INVALID_HANDLE_VALUE          As Long = (-1)              '無効なFile Handle値
Private Const FILE_ATTRIBUTE_HIDDEN         As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM         As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10               'フォルダ
Private Const GENERIC_READ                  As Long = &H80000000
Private Const GENERIC_WRITE                 As Long = &H40000000
Private Const FILE_SHARE_READ               As Long = &H1
Private Const FILE_ATTRIBUTE_NORMAL         As Long = &H80
Private Const OPEN_EXISTING                 As Long = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS    As Long = &H2000000

Private Const FindExInfoStandard As Long = 0&
Private Const FindExInfoBasic As Long = 1&              'Win7以降
Private Const FindExSearchNameMatch As Long = 0
Private Const FindExSearchLimitToDirectories = 1&
Private Const FIND_FIRST_EX_LARGE_FETCH As Long = 2&    'Win7以降

Private FSO As Object

' FileTime 構造体
Private Type FILETIME
    LowDateTime     As Long
    HighDateTime    As Long
End Type

'WIN32_FIND_DATA構造体(ディレクトリエントリ(ファイル情報))
Private Type WIN32_FIND_DATA
    dwFileAttributes            As Long             'ファイル属性
    ftCreationTime              As FILETIME         '作成日時
    ftLastAccessTime            As FILETIME         '最終アクセス日時
    ftLastWriteTime             As FILETIME         '最終更新日時
    nFileSizeHigh               As Long             'ファイルサイズの上位32bit値
    nFileSizeLow                As Long             'ファイルサイズの下位32bit値
    dwReserved0                 As Long             '予約(現状なし)
    dwReserved1                 As Long             '予約(現状なし)
    cFileName(MAX_PATH)         As Byte             'ロングファイル名
    cAlternate(MAX_PATH_SHORT)  As Byte             'ショートファイル名(8+3文字)
    dwFileType                  As Long
    dwCreatorType               As Long
    wFinderFlags                As Long
End Type

' SystemTime 構造体
Private Type SystemTime
    Year            As Integer
    Month           As Integer
    DayOfWeek       As Integer
    Day             As Integer
    Hour            As Integer
    Minute          As Integer
    Second          As Integer
    Milliseconds    As Integer
End Type

Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub
Private Sub Class_Terminate()
    Set FSO = Nothing
End Sub
'--------------------------------------------------------------
'　ファイル存在チェック
'--------------------------------------------------------------
Public Function FileExists(ByVal strFile As String) As Boolean
 
    With FSO
        FileExists = .FileExists(strFile)
    End With

End Function
'--------------------------------------------------------------
'　ファイルオブジェクト取得
'--------------------------------------------------------------
Public Function GetFile(ByVal strFile As String) As Object
 
    With FSO
        Set GetFile = .GetFile(strFile)
    End With

End Function
'--------------------------------------------------------------
'　フォルダーオブジェクト取得
'--------------------------------------------------------------
Public Function GetFolder(ByVal strFolder As String) As Object
 
    With FSO
        Set GetFolder = .GetFolder(strFolder)
    End With

End Function
'--------------------------------------------------------------
'　フォルダ存在チェック
'--------------------------------------------------------------
Public Function FolderExists(ByVal strFile As String) As Boolean
 
    With FSO
        FolderExists = .FolderExists(strFile)
    End With

End Function
'--------------------------------------------------------------
'　テンポラリフォルダ取得
'--------------------------------------------------------------
Public Property Get TempFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    TempFolder = ""
    
    With FSO
    
        strFolder = AppDataFolder & "Temp"
        
        If .FolderExists(strFolder) Then
        Else
            .CreateFolder strFolder
        End If
        
        TempFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Property
'--------------------------------------------------------------
'　アプリケーションフォルダ取得
'--------------------------------------------------------------
Public Property Get AppDataFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    AppDataFolder = ""
    
    With FSO
    
        strFolder = .BuildPath(CreateObject("Wscript.Shell").SpecialFolders("AppData"), Title)
        
        If .FolderExists(strFolder) Then
        Else
            .CreateFolder strFolder
        End If
        
        AppDataFolder = .BuildPath(strFolder, "\")
        
    End With

End Property
'--------------------------------------------------------------
'　ファイル検索
'--------------------------------------------------------------
Public Sub FileSearch(ByVal strPath As String, _
                      ByVal varPatterns As Variant, _
                      ByRef col As Collection, _
             Optional ByVal SubDir As Boolean = True, _
             Optional ByVal NaturalSort = True)

    Dim objfld As Object
    Dim objfl As Object
    Dim objSub As Object
    Dim v As Variant

    With FSO

        Set objfld = .GetFolder(strPath)

        'ファイル名取得
        For Each objfl In objfld.files
            For Each v In varPatterns
                If LCase(objfl.Name) Like LCase(v) And Left(objfl.Name, 2) <> "~$" Then
                    Append col, .BuildPath(objfl.ParentFolder.Path, objfl.Name), NaturalSort
                    Exit For
                End If
            Next
            DoEvents
        Next

        'サブフォルダ検索
        If SubDir Then
            For Each objSub In objfld.SubFolders
                FileSearch objSub.Path, varPatterns, col, SubDir, NaturalSort
                DoEvents
            Next
        End If
    End With
End Sub
'--------------------------------------------------------------------------------------------------------------------
'　ファイル検索(フィルタ高速版)
'
'  FileSystemObject は便利だが、フィルタがかけられず、全部のファイルを列挙するためファイル検索が遅い問題がある。
'  VBA.DIRは3桁以上の拡張子非対応、UNC非対応のため、Winddows API にて実装
'  自然数ソートを追加。
'  SubDir = True(-1):制限なし False(0):指定フォルダのみ 階数:指定階数まで検索
'--------------------------------------------------------------------------------------------------------------------
Public Sub FileSearchEx(ByVal strPath As String, _
                        ByVal varPatterns As Variant, _
                        ByRef col As Collection, _
               Optional ByVal SubDir As Long = 0, _
               Optional ByVal ReturnFind As Boolean = False, _
               Optional ByVal NaturalSort As Boolean = True, _
               Optional ByVal varExcludedFolder As Variant = "")

'    Dim PL As PairLogger: Set PL = Constructor(New PairLogger, TypeName(Me) & ".FileSearchEx")

    Dim v As Variant
    Dim udtWin32FindData        As WIN32_FIND_DATA      '[WIN32_FIND_DATA]構造体(検索結果)
    Dim hResult As LongPtr
    Dim strSeachFullPath As String
    Dim strFile As String
    Dim strFindFileName As String               '検索結果ファイル名
    
    Static lngLevel As Long
    
    '指定レベル以下の場合、検索しない
    Select Case True
        Case SubDir = -1
        Case SubDir < lngLevel
            Exit Sub
    End Select
    
    'Collectionにインスタンスが無い場合設定
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    '配列じゃない場合、配列に変換
    If Not IsArray(varPatterns) Then
        varPatterns = Array(varPatterns)
    End If

    For Each v In varPatterns
    
        '検索フルパス名を生成
        If strPath Like "\\*" Then
            strSeachFullPath = "\\?\UNC\" & Mid$(strPath, 3)
        Else
            strSeachFullPath = "\\?\" & strPath
        End If
        
        strSeachFullPath = FileIO.BuildPath(strSeachFullPath, v)
        
        '文字列に一致するファイルを検索し、WIN32_FIND_DATA構造体に値を代入
        hResult = FindFirstFileExW(StrPtr(strSeachFullPath), FindExInfoBasic, udtWin32FindData, FindExSearchNameMatch, 0, FIND_FIRST_EX_LARGE_FETCH)
        
        'ファイル有の場合
        If hResult <> INVALID_HANDLE_VALUE Then

            Do
                strFindFileName = CStr(udtWin32FindData.cFileName)
                strFile = Left$(strFindFileName, InStr(strFindFileName, vbNullChar) - 1)
                
                Select Case True
                    Case strFile = "."
                    Case strFile = ".."
                    Case Not strFile Like v 'FindFirstFile/FindNextFileで拡張子４桁以上は無視されてしまうため別途Likeで判定する。
                    Case Left$(strFile, 2) = "~$"
                    Case udtWin32FindData.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN
                    Case udtWin32FindData.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM
                    Case udtWin32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
                    Case Else
                        
                        '上記以外を追加
                        Append col, FileIO.BuildPath(strPath, strFile), NaturalSort
                        
                        '即リターン
                        If ReturnFind Then
                            hResult = FindClose(hResult)
                            Exit Sub
                        End If
                            
                End Select
            
            Loop While FindNextFileW(hResult, udtWin32FindData)
        
             'ファイルハンドルをクローズ
            hResult = FindClose(hResult)
        
        End If
    Next
    
    'サブフォルダ検索
    Dim objfld As Object
    Dim objSub As Object

    If SubDir Then
    
        If Not IsArray(varExcludedFolder) Then
            varExcludedFolder = Array(varExcludedFolder)
        End If

        Set objfld = FSO.GetFolder(strPath)
        For Each objSub In objfld.SubFolders
        
            '除外フォルダループ
            For Each v In varExcludedFolder
            
                '一致しないものについて取得
                If v <> FileIO.GetFileName(objSub.Path) Then
            
                    '再帰呼び出し
                    lngLevel = lngLevel + 1
                    FileSearchEx objSub.Path, varPatterns, col, SubDir, ReturnFind, NaturalSort, varExcludedFolder
                    lngLevel = lngLevel - 1
        
                    '即リターン
                    If ReturnFind And col.Count > 0 Then
                        Exit Sub
                    End If
                    
                End If
            Next
            DoEvents
        Next

    End If

End Sub
'自然数挿入ソート
Private Sub Append(ByRef col As Collection, _
                   ByVal strFile As String, _
                   ByVal NaturalSort As Boolean)

    Dim i As Long
    
    If NaturalSort Then
        If col.Count = 0 Then
            col.Add strFile
        Else
            For i = col.Count To 1 Step -1

                'ファイルの方が大きかった場合、その後に挿入。
                If StrCmpLogicalW(StrPtr(strFile), StrPtr(col(i))) >= 0 Then
                    col.Add strFile, , , i
                    Exit Sub
                End If

            Next
            col.Add strFile, , 1
        End If
    Else
        col.Add strFile
    End If

End Sub

'--------------------------------------------------------------
'　ファイル検索(フィルタ高速版)
'  指定のファイルがファイルが見つかったら即リターンします。
'--------------------------------------------------------------
Public Function SearchTreeForFile(ByVal strPath As String, ByVal strFile As String) As String

    Dim strBuffer As String
    
    strBuffer = String$(MAX_PATH, vbNullChar)
    
    SearchTreeForFile = ""

    If SearchTreeForFileW(StrPtr(strPath), StrPtr(strFile), StrPtr(strBuffer)) Then
        SearchTreeForFile = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If

End Function
'--------------------------------------------------------------
'　ファイルパス結合
'--------------------------------------------------------------
Public Function BuildPath(ByVal strPath As String, ByVal strFile As String) As String
    
    With FSO
        BuildPath = .BuildPath(strPath, strFile)
    End With

End Function
'--------------------------------------------------------------
'　フォルダ削除
'  指定フォルダを以下を含む、フォルダ自体を削除
'--------------------------------------------------------------
Public Sub DeleteFolder(ByVal strPath As String)
    
    With FSO
        .DeleteFolder strPath, True
    End With

End Sub
'--------------------------------------------------------------
'　ファイル削除
'--------------------------------------------------------------
Public Sub DeleteFile(ByVal strPath As String)
    
    With FSO
        .DeleteFile strPath, True
    End With

End Sub
'--------------------------------------------------------------
'　指定フォルダ以下ファイル＆フォルダ削除
'  指定フォルダは消さずにそれ以下を消すことに注意
'  ワークフォルダのクリアに向く。
'--------------------------------------------------------------
Public Sub ClearFolder(ByVal strPath As String)
    
    Dim fld As Object
    Dim f As Object
    
    With FSO
    
        Set fld = .GetFolder(strPath)
        
        For Each f In fld.files
            .DeleteFile .BuildPath(strPath, f.Name), True
        Next
        
        For Each f In fld.SubFolders
            .DeleteFolder .BuildPath(strPath, f.Name), True
        Next
        
    End With

End Sub

'--------------------------------------------------------------
'　ファイル名取得
'--------------------------------------------------------------
Public Function GetFileName(ByVal strFile As String) As String

    With FSO

        GetFileName = .GetFileName(strFile)
        
    End With
    
'    Dim lngCnt As Long
'    Dim lngMax As Long
'    Dim strResult As String
'
'    strResult = strPath
'
'    lngMax = Len(strPath)
'
'    For lngCnt = lngMax To 1 Step -1
'
'        Select Case Mid$(strPath, lngCnt, 1)
'            Case "\", "/"
'                If lngCnt = lngMax Then
'                Else
'                    strResult = Mid$(strPath, lngCnt + 1)
'                End If
'                Exit For
'        End Select
'
'    Next
'
'    rlxGetFullpathFromFileName = strResult

End Function
'--------------------------------------------------------------
'　拡張子取得
'--------------------------------------------------------------
Public Function GetExtensionName(ByVal strFile As String) As String

    With FSO

        GetExtensionName = .GetExtensionName(strFile)
        
    End With
    
'   Dim lngCnt As Long
'    Dim lngMax As Long
'    Dim strResult As String
'
'    strResult = strPath
'
'    lngMax = Len(strPath)
'
'    For lngCnt = lngMax To 1 Step -1
'
'        If Mid$(strPath, lngCnt, 1) = "." Then
'            If lngCnt > 1 Then
'                strResult = Mid$(strPath, 1, lngCnt - 1)
'                Exit For
'            End If
'        End If
'
'    Next
'
'    rlxGetFullpathFromExt = strResult

End Function
'--------------------------------------------------------------
'　ファイル名取得(拡張子抜き)
'--------------------------------------------------------------
Public Function GetBaseName(ByVal strFile As String) As String

    With FSO
        GetBaseName = .GetBaseName(strFile)
    End With

End Function
'--------------------------------------------------------------
'　パス情報取得
'--------------------------------------------------------------
Public Function GetParentFolderName(ByVal strFile As String) As String

    With FSO

        GetParentFolderName = .GetParentFolderName(strFile)
        
    End With
    
'    Dim lngCnt As Long
'    Dim lngMax As Long
'    Dim strResult As String
'
'    strResult = strPath
'
'    lngMax = Len(strPath)
'
'    For lngCnt = lngMax To 1 Step -1
'
'        Select Case Mid$(strPath, lngCnt, 1)
'            Case "\", "/"
'                If lngCnt > 1 Then
'                    strResult = Mid$(strPath, 1, lngCnt - 1)
'                    Exit For
'                End If
'        End Select
'
'    Next
'
'    rlxGetFullpathFromPathName = strResult

End Function
'--------------------------------------------------------------
'  フォルダの作成（再帰）
'--------------------------------------------------------------
Public Sub CreateFolder(ByVal strPath As String)

    With FSO
    
        If .FolderExists(strPath) Then
            Exit Sub
        End If
    
        Call CreateFolder(.GetParentFolderName(strPath))
        
        .CreateFolder strPath
    
    End With

End Sub
'--------------------------------------------------------------
'  マイドキュメントフォルダ移動
'--------------------------------------------------------------
Public Sub SetMyDocument()
    On Error Resume Next
    ChDir CreateObject("Wscript.Shell").SpecialFolders("MyDocuments")
End Sub
'--------------------------------------------------------------
'　ドライブ名→UNC名変換
'　ドライブ名(J:等)を指定。エラーの場合ドライブ名をそのまま返却
'--------------------------------------------------------------
Public Function DriveToUNC(ByVal strPath As String) As String

    Dim lStatus As Long
    Dim lpLocalName As String
    Dim lpRemoteName As String
    Dim lpnLength As Long
    
    Const NO_ERROR As Long = 0
    
    'デフォルトでパスをセット
    DriveToUNC = strPath
    
    If InStr(strPath, ":") = 2 Then
        lpLocalName = Left$(strPath, 2)
    Else
        'ドライブ情報が含まれない。
        Exit Function
    End If

    lpnLength = MAX_PATH
    
    lpRemoteName = String$(MAX_PATH, vbNullChar)
    
    lStatus& = WNetGetConnectionW(StrPtr(lpLocalName), StrPtr(lpRemoteName), lpnLength)
    
    If lStatus& = NO_ERROR Then
        DriveToUNC = Left$(lpRemoteName, InStr(lpRemoteName, vbNullChar) - 1) & Mid$(strPath, 3)
    End If

End Function
'--------------------------------------------------------------
'  フォルダを開く
'--------------------------------------------------------------
Public Sub OpenFolder(ByVal strFolder As String)

    With CreateObject("WScript.Shell")
        .Run ("""" & strFolder & """")
    End With
    
End Sub
'--------------------------------------------------------------
'  ファイルをエクスプローラーで開いて選択
'--------------------------------------------------------------
'  初心者忘備録
'  https://www.ka-net.org/blog/?p=9180
'  指定したファイルをエクスプローラーで開いて選択するVBAマクロ
'--------------------------------------------------------------
Public Sub OpenFileOnExplorer(ByVal strFile As String)

    '指定したファイルをエクスプローラーで開いて選択する
    With FSO
        If .FileExists(strFile) Then
            VBA.Shell "EXPLORER.EXE /select,""" & strFile & """", vbNormalFocus
        End If
    End With

End Sub
''--------------------------------------------------------------
''  ファイルコピー
''--------------------------------------------------------------
'Public Sub CopyFile(ByVal source As String, ByVal destination As String)
'
'    With FileSystemObject
'        .CopyFile source, destination, True
'    End With
'
'End Sub
'--------------------------------------------------------------
'  ファイルコピー(タイムスタンプ維持＆読み取り専用の上書きコピー対応＆対象ファイル無でもエラーにならない)
'--------------------------------------------------------------
Public Sub CopyFile(ByVal source As String, _
                    ByVal destination As String, _
           Optional ByVal keepTimestamp As Boolean = False)
    
    Dim dateCreated As Date
    Dim dateLastModified As Date
    Dim dateLastAccessed As Date
    Dim udtWin32FindData        As WIN32_FIND_DATA      '[WIN32_FIND_DATA]構造体(検索結果)
    Dim hResult As LongPtr
    Dim strSeachFullPath As String
    Dim strFile As String

    With FSO
        
        '検索フルパス名を生成
        If source Like "\\*" Then
            strSeachFullPath = "\\?\UNC\" & Mid$(source, 3)
        Else
            strSeachFullPath = "\\?\" & source
        End If
        
        '文字列に一致するファイルを検索し、WIN32_FIND_DATA構造体に値を代入
        hResult = FindFirstFileW(StrPtr(strSeachFullPath), udtWin32FindData)
        'ファイル無の場合
        If hResult = INVALID_HANDLE_VALUE Then
            Exit Sub
        End If
        
        Do
            Dim strFindFileName As String
            
            strFindFileName = CStr(udtWin32FindData.cFileName)
            strFile = Left$(strFindFileName, InStr(strFindFileName, vbNullChar) - 1)
            
            'ファイルの属性がディレクトリの場合、パス
            If strFile = "" Or strFile = "." Or strFile = ".." Then
                GoTo pass
            End If
                
            Select Case True
                Case udtWin32FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
                    GoTo pass
            End Select
                
            Dim strFullSource As String
            strFullSource = FileIO.BuildPath(FileIO.GetParentFolderName(source), strFile)
            
            'タイムスタンプ取得
            If keepTimestamp Then
                dateCreated = .GetFile(strFullSource).dateCreated
                dateLastModified = .GetFile(strFullSource).dateLastModified
                dateLastAccessed = .GetFile(strFullSource).dateLastAccessed
            End If
            
            'コピー先がフォルダ指定の場合、コピー先ファイル名を作成
            Dim strFullDest As String
            If FileIO.FolderExists(destination) Then
                strFullDest = FileIO.BuildPath(destination, strFile)
            Else
                strFullDest = destination
            End If
            
            '読み取り専用ファイル対策。事前にファイルを削除
            If FileIO.FileExists(strFullDest) Then
                .DeleteFile strFullDest, True
            End If
            
            .CopyFile strFullSource, strFullDest, True
            
            'タイムスタンプ書き込み
            If keepTimestamp Then
                
                Dim lngAttr As Long
                Dim objFile As Object
                
                Set objFile = .GetFile(strFullDest)
                lngAttr = objFile.Attributes
                
                '標準ファイルに変更（読み取り専用対策）
                objFile.Attributes = 0
                
                'タイムスタンプ書き込み
                FileIO.SetCreationTime strFullDest, dateCreated
                FileIO.SetLastWriteTime strFullDest, dateLastModified
                FileIO.SetLastAccessTime strFullDest, dateLastAccessed
                
                '属性を元に戻す
                objFile.Attributes = lngAttr
            
            End If
    
pass:
        Loop While FindNextFileW(hResult, udtWin32FindData)
    
         'ファイルハンドルをクローズ
        hResult = FindClose(hResult)
    
    End With

End Sub
'--------------------------------------------------------------
'  ファイル／フォルダの作成日時設定
'--------------------------------------------------------------
Public Sub SetCreationTime(ByVal stFilePath As String, ByVal dtCreateTime As Date)
    
    Dim cFileHandle As LongPtr
    Dim tFileTime As FILETIME
    Dim tNullable As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtCreateTime)
        Call SetFileTime(cFileHandle, tFileTime, tNullable, tNullable)
        Call CloseHandle(cFileHandle)
    End If
    
End Sub
'--------------------------------------------------------------
'  ファイル／フォルダの更新日時設定
'--------------------------------------------------------------
Public Sub SetLastWriteTime(ByVal stFilePath As String, ByVal dtUpdateTime As Date)
    
    Dim cFileHandle As LongPtr
    Dim tFileTime As FILETIME
    Dim tNullable As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtUpdateTime)
        Call SetFileTime(cFileHandle, tNullable, tNullable, tFileTime)
        Call CloseHandle(cFileHandle)
    End If

End Sub
'--------------------------------------------------------------
'  ファイル／フォルダのアクセス日時設定
'--------------------------------------------------------------
Public Sub SetLastAccessTime(ByVal stFilePath As String, ByVal dtAccessTime As Date)

    Dim cFileHandle As LongPtr
    Dim tNullable As FILETIME
    Dim tFileTime As FILETIME
    
    cFileHandle = GetFileHandle(stFilePath)
    If cFileHandle <> 0 Then
        tFileTime = GetFileTime(dtAccessTime)
        Call SetFileTime(cFileHandle, tNullable, tFileTime, tNullable)
        Call CloseHandle(cFileHandle)
    End If
    
End Sub

' FileTime を取得する
Private Function GetFileTime(ByVal dtSetting As Date) As FILETIME

    Dim tSystemTime As SystemTime
    
    With tSystemTime
        .Year = Year(dtSetting)
        .Month = Month(dtSetting)
        .DayOfWeek = Weekday(dtSetting)
        .Day = Day(dtSetting)
        .Hour = Hour(dtSetting)
        .Minute = Minute(dtSetting)
        .Second = Second(dtSetting)
    End With
    
    Dim tLocalTime As FILETIME
    Call SystemTimeToFileTime(tSystemTime, tLocalTime)
    
    Dim tFileTime As FILETIME
    Call LocalFileTimeToFileTime(tLocalTime, tFileTime)
    
    GetFileTime = tFileTime
    
End Function

' ファイルのハンドルを取得する
Private Function GetFileHandle(ByVal stFilePath As String) As LongPtr

'フォルダのハンドルの取得に対応
    'GetFileHandle = CreateFile(StrPtr(stFilePath), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    GetFileHandle = CreateFileW(StrPtr(stFilePath), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)

End Function

'--------------------------------------------------------------
' 元フォルダと同じ構成のフォルダを作成する。
'--------------------------------------------------------------
Public Sub CreateFolderImage(ByVal strSource As String, ByVal strDest As String)

    Dim strFolder As String
    Dim objfld As Object
    Dim objSub As Object

    With FSO

        Set objfld = .GetFolder(strSource)
    
        For Each objSub In objfld.SubFolders
            
            '作成側フォルダ名を生成
            strFolder = FileIO.BuildPath(strDest, objSub.Name)
            
            If Not FileIO.FolderExists(strFolder) Then
                
                'フォルダ作成
                FileIO.CreateFolder strFolder
                
                Dim dt As Date
                dt = .GetFolder(objSub.Path).dateCreated
                FileIO.SetCreationTime strFolder, dt
                
                dt = .GetFolder(objSub.Path).dateLastModified
                FileIO.SetLastWriteTime strFolder, dt
            
                dt = .GetFolder(objSub.Path).dateLastAccessed
                FileIO.SetLastAccessTime strFolder, dt
            
            End If
            
            'サブフォルダを再帰にて作成
            CreateFolderImage objSub.Path, strFolder
            
            DoEvents
        Next
    
    End With

End Sub


'アプリケーション名取得
Private Property Get Title() As String
    Title = ThisWorkbook.BuiltinDocumentProperties("Title").Value
End Property
'--------------------------------------------------------------
'　ファイルを削除するのではなく内容のクリア
'--------------------------------------------------------------
Sub TruncateFile(ByVal strFile As String)

    Dim fp As Integer

    '出力ファイルのクリア
    fp = FreeFile()
    Open strFile For Output As fp
    Close fp

End Sub
