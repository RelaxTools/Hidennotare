VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileIO"
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
'>### FileIO クラス
'>
'>**Remarks**
'>
'>CreateObject しなくても使える FileSystemObject<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
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
'--------------------------------------------------------------
' Inisialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
End Sub
'--------------------------------------------------------------
' Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set FSO = Nothing
End Sub
'>---
'>#### FileExists(FileIO)
'>
'>**Syntax**
'>
'>```
'>flg = FileIO.FileExists(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|存在チェックを行うファイルを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|存在する場合:True / 存在しない場合:False|
'>
'>
'>**Remarks**
'>
'>ファイル存在チェック<br>
'>指定されたファイルの存在チェックを行う
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.FolderExists
'>
Public Function FileExists(ByVal strFile As String) As Boolean
 
    With FSO
        FileExists = .FileExists(strFile)
    End With

End Function
'>---
'>#### FolderExists(FileIO)
'>
'>**Syntax**
'>
'>```
'>boolean = FileIO.FolderExists(strFolder)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFolder|必須|String|存在チェックを行うフォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|flg|Boolean|存在する場合:True / 存在しない場合:False|
'>
'>**Remarks**
'>
'>フォルダ存在チェック<br>
'>指定されたフォルダルの存在チェックを行う
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.FileExists
'>
Public Function FolderExists(ByVal strFolder As String) As Boolean
 
    With FSO
        FolderExists = .FolderExists(strFolder)
    End With

End Function
'>---
'>#### GetFile(FileIO)
'>
'>**Syntax**
'>
'>```
'>Set ret = FileIO.GetFile(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|ファイルを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|Scripting.FileSystemObject の FileObject|
'>
'>**Remarks**
'>
'>ファイルオブジェクト取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetFolder
'>
Public Function GetFile(ByVal strFile As String) As Object
 
    With FSO
        Set GetFile = .GetFile(strFile)
    End With

End Function
'>---
'>#### GetFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>Set ret = FileIO.GetFolder(strFolder)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFolder|必須|String|フォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|Object|Scripting.FileSystemObject の FolderObject|
'>
'>**Remarks**
'>
'>フォルダオブジェクト取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetFile
'>
Public Function GetFolder(ByVal strFolder As String) As Object
 
    With FSO
        Set GetFolder = .GetFolder(strFolder)
    End With

End Function
'>---
'>#### TempFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.TempFolder()
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
'>|ret|String|テンポラリフォルダのパス|
'>
'>**Remarks**
'>
'>テンポラリフォルダ取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.AppDataFolder
'>
Public Property Get TempFolder() As String

    TempFolder = FSO.GetSpecialFolder(2).Path

End Property
'>---
'>#### PrivateTempFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.PrivateTempFolder()
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
'>|ret|String|テンポラリフォルダのパス|
'>
'>**Remarks**
'>
'>テンポラリフォルダ取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.AppDataFolder
'>
Public Property Get PrivateTempFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    PrivateTempFolder = ""
    
    With FSO
    
        strFolder = AppDataFolder & "Temp"
        
        If .FolderExists(strFolder) Then
        Else
            .CreateFolder strFolder
        End If
        
        PrivateTempFolder = .BuildPath(strFolder, "\")
        
    End With
    

End Property
'>---
'>#### GetTempName(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.GetTempName()
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
'>|ret|String|テンポラリファイル名|
'>
'>**Remarks**
'>
'>テンポラリファイル名取得
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.TempFolder
'>
Public Property Get GetTempName() As String
    
    With FSO
        
        GetTempName = .GetTempName
        
    End With
    

End Property
'>---
'>#### AppDataFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>folder = FileIO.AppDataFolder()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>アプリケーションフォルダのパス
'>
'>**Remarks**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|アプリケーションフォルダ取得|
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.TempFolder
'>
Public Property Get AppDataFolder() As String

    On Error Resume Next
    
    Dim strFolder As String
    
    AppDataFolder = ""
    
    With FSO
    
        strFolder = .BuildPath(VBA.CreateObject("Wscript.Shell").SpecialFolders("AppData"), Title)
        
        If .FolderExists(strFolder) Then
        Else
            .CreateFolder strFolder
        End If
        
        AppDataFolder = .BuildPath(strFolder, "\")
        
    End With

End Property
'>---
'>#### FileSearch(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.FileSearch strPath,
'>                  varPatterns,
'>                  Col,
'>                  [SubDir = True],
'>                  [NaturalSort = True]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|検索するフォルダを指定|
'>|varPatterns|必須|Variant|ファイルのパターン|
'>|Col|必須|Collection|検索したファイル一覧を返却する。|
'>|SubDir|省略可|Boolean|Trueの場合サブフォルダを検索する。|
'>|NaturalSort|省略可|NaturalSort|Trueの場合、自然数ソートを行う。省略時True|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル検索<br>
'>FileSearchExの方が高速で多機能です。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.FileSearch
'>* FileIO.FileSearchEx
'>* FileIO.SearchTreeForFile
'>
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
'>---
'>#### FileSearchEx(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.FileSearchEx strPath,
'>                    varPatterns,
'>                    Col,
'>                    [SubDir = True],
'>                    [ReturnFind = False],
'>                    [NaturalSort = True],
'>                    [varExcludedFolder = ""]
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|検索するフォルダを指定|
'>|varPatterns|必須|Variant|ファイルのパターンを指定。複数の場合には配列で指定。ワイルドカードが使用可能。|
'>|Col|必須|Collection|検索したファイル一覧を返却する。|
'>|SubDir|省略可|Boolean|True(-1):制限なし/False(0):指定フォルダのみ/階数:指定階数まで検索|
'>|ReturnFind|省略可|Boolean|Trueの場合、ファイルが見つかった時点で即リターン。省略時True|
'>|NaturalSort|省略可|Boolean|Trueの場合、自然数ソートを行う。省略時True|
'>|varExcludedFolder|省略可|除外フォルダを指定。複数の場合には配列で指定。|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル一覧取得(フィルタ高速版)<br>
'>FileSystemObject は便利だが、フィルタがかけられず、全部のファイルを列挙するためファイル検索が遅い問題がある。<br>
'>VBA.DIRは3桁以上の拡張子非対応、UNC非対応のため、Winddows API にて実装<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.FileSearch
'>* FileIO.SearchTreeForFile
'>
Public Sub FileSearchEx(ByVal strPath As String, _
                        ByVal varPatterns As Variant, _
                        ByRef col As Collection, _
               Optional ByVal SubDir As Long = 0, _
               Optional ByVal ReturnFind As Boolean = False, _
               Optional ByVal NaturalSort As Boolean = True, _
               Optional ByVal varExcludedFolder As Variant = "")

'    Dim PL As PairLogger: Set PL = PairLogger.CreateObject(TypeName(Me) & ".FileSearchEx")

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
'>---
'>#### SearchTreeForFile(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.SearchTreeForFile(strPath, strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|検索するフォルダを指定|
'>|strFile|必須|String|ファイルのパターンを指定。ワイルドカードが使用可能。|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|検索したファイル名|
'>
'>**Remarks**
'>
'>ファイル検索<br>
'>指定のファイルがファイルが見つかったら即リターンします。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.FileSearch
'>* FileIO.FileSearchEx
'>* FileIO.SearchTreeForFile
'>
Public Function SearchTreeForFile(ByVal strPath As String, ByVal strFile As String) As String

    Dim strBuffer As String
    
    strBuffer = String$(MAX_PATH, vbNullChar)
    
    SearchTreeForFile = ""

    If SearchTreeForFileW(StrPtr(strPath), StrPtr(strFile), StrPtr(strBuffer)) Then
        SearchTreeForFile = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If

End Function
'>---
'>#### BuildPath(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.BuildPath(strPath, strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|結合するフォルダを指定|
'>|strFile|必須|String|結合ファイルを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|結合したファイル名|
'>
'>**Remarks**
'>
'>ファイルパス結合<br>
'>指定のファイルがファイルが見つかったら即リターンします。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Function BuildPath(ByVal strPath As String, ByVal strFile As String) As String
    
    With FSO
        BuildPath = .BuildPath(strPath, strFile)
    End With

End Function
'>---
'>#### DeleteFile(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.DeleteFile(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|削除するファイルを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル削除<br>
'>指定ファイルを削除<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.TruncateFile
'>
Public Sub DeleteFile(ByVal strPath As String)
    
    With FSO
        .DeleteFile strPath, True
    End With

End Sub
'>---
'>#### TruncateFile(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.TruncateFile(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|クリアするファイルを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル内容のクリア<br>
'>ファイルを削除するのではなく内容のクリア<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.DeleteFile
'>
Sub TruncateFile(ByVal strFile As String)

    Dim FP As Integer

    '出力ファイルのクリア
    FP = FreeFile()
    Open strFile For Output As FP
    Close FP

End Sub
'>---
'>#### DeleteFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.DeleteFolder(strPath)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|削除するフォルダを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>フォルダ削除<br>
'>指定フォルダを以下を含む、フォルダ自体を削除<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.ClearFolder
'>
Public Sub DeleteFolder(ByVal strPath As String)
    
    With FSO
        .DeleteFolder strPath, True
    End With

End Sub
'>---
'>#### ClearFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.ClearFolder(strPath)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|クリアするフォルダを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>指定フォルダ以下ファイル＆フォルダ削除<br>
'>指定フォルダは消さずにそれ以下を消すことに注意<br>
'>ワークフォルダのクリアに向く<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.DeleteFolder
'>
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
'>---
'>#### ClearPrivateTempFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.ClearPrivateTempFolder()
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>テンポラリフォルダ以下ファイル＆フォルダ削除<br>
'>テンポラリフォルダは消さずにそれ以下を消すことに注意<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.PrivateTempFolder
'>
Public Sub ClearPrivateTempFolder()
    
    Dim fld As Object
    Dim f As Object
    Dim strPath As String
    
    strPath = FileIO.TempFolder
    
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

'>---
'>#### CreateFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.CreateFolder(strPath)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|作成するフォルダを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>フォルダ作成<br>
'>上位フォルダが無い場合、再帰で作成する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub CreateFolder(ByVal strPath As String)

    With FSO
    
        If .FolderExists(strPath) Then
            Exit Sub
        End If
    
        Call CreateFolder(.GetParentFolderName(strPath))
        
        .CreateFolder strPath
    
    End With

End Sub
'>---
'>#### GetFileName(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.GetFileName(strPath)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|ファイルを含むフォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|ファイル名のみを返却|
'>
'>**Remarks**
'>
'>ファイル名取得<br>
'>パス情報からファイル名を取得<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetExtensionName
'>* FileIO.GetBaseName
'>* FileIO.GetParentFolderName
'>
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
'>---
'>#### GetExtensionName(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.GetExtensionName(strPath)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|ファイルを含むフォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|拡張子のみを返却|
'>
'>**Remarks**
'>
'>拡張子取得<br>
'>パス情報又はファイルから拡張子を取得<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetFileName
'>* FileIO.GetBaseName
'>* FileIO.GetParentFolderName
'>
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
'>---
'>#### GetBaseName(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.GetBaseName(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|ファイルを含むフォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|拡張子を除いたファイル名のみを返却|
'>
'>
'>**Remarks**
'>
'>ファイル名取得（拡張子除く）<br>
'>拡張子を除いたファイル名のみを取得<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetFileName
'>* FileIO.GetExtensionName
'>* FileIO.GetParentFolderName
'>
Public Function GetBaseName(ByVal strFile As String) As String

    With FSO
        GetBaseName = .GetBaseName(strFile)
    End With

End Function
'>---
'>#### GetParentFolderName(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.GetParentFolderName(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strPath|必須|String|ファイルを含むフォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|上位パスのみを返却|
'>
'>**Remarks**
'>
'>パス情報取得<br>
'>指定ファイルの上位パスを返却する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.GetFileName
'>* FileIO.GetExtensionName
'>* FileIO.GetBaseName
'>
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
'>---
'>#### SetMyDocument(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.SetMyDocument(strFile)
'>```
'>
'>**Parameters**
'>
'>* None
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>マイドキュメントフォルダ移動<br>
'>カレントフォルダをマイドキュメントに移動する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub SetMyDocument()
    On Error Resume Next
    ChDir VBA.CreateObject("Wscript.Shell").SpecialFolders("MyDocuments")
End Sub
'>---
'>#### DriveToUNC(FileIO)
'>
'>**Syntax**
'>
'>```
'>ret = FileIO.DriveToUNC(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|ドライブを含むファイル／フォルダを指定|
'>
'>**Return Value**
'>
'>|Name|Data type|Description|
'>---|---|---
'>|ret|String|UNCパスを返却|
'>
'>**Remarks**
'>
'>ドライブ名→UNC名変換<br>
'>ドライブ名(J:等)を指定。エラーの場合ドライブ名をそのまま返却<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
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
'>---
'>#### OpenFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.OpenFolder(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFolder|必須|String|開くフォルダを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>フォルダを開く<br>
'>指定フォルダをエクスプローラで開く<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.OpenFileOnExplorer
'>
Public Sub OpenFolder(ByVal strFolder As String)

    With VBA.CreateObject("WScript.Shell")
        .Run ("""" & strFolder & """")
    End With
    
End Sub
'>---
'>#### OpenFileOnExplorer(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.OpenFileOnExplorer(strFile)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|strFile|必須|String|開くファイルを指定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイルをエクスプローラーで開いて選択<br>
'>初心者忘備録<br>
'>https://www.ka-net.org/blog/?p=9180<br>
'>指定したファイルをエクスプローラーで開いて選択するVBAマクロ<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.OpenFolder
'>
Public Sub OpenFileOnExplorer(ByVal strFile As String)

    '指定したファイルをエクスプローラーで開いて選択する
    With FSO
        If .FileExists(strFile) Then
            VBA.Shell "EXPLORER.EXE /select,""" & strFile & """", vbNormalFocus
        End If
    End With

End Sub
'>---
'>#### CopyFile(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.CopyFile(source,
'>                destination,
'>                [keepTimestamp = False])
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|source|必須|String|コピー元|
'>|destination|必須|コピー先|
'>|keepTimestamp|必須|True:タイムスタンプを維持する/False:維持しない。デフォルトFalse|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイルコピー<br>
'>タイムスタンプ維持＆読み取り専用の上書きコピー対応＆対象ファイル無でもエラーにならない<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Public Sub CopyFile(ByVal Source As String, _
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
        If Source Like "\\*" Then
            strSeachFullPath = "\\?\UNC\" & Mid$(Source, 3)
        Else
            strSeachFullPath = "\\?\" & Source
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
            strFullSource = FileIO.BuildPath(FileIO.GetParentFolderName(Source), strFile)
            
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
'>---
'>#### SetCreationTime(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.SetCreationTime(stFilePath, dtCreateTime)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|stFilePath|必須|String|作成日付を設定するファイルを指定|
'>|dtCreateTime|必須|作成日付を設定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル／フォルダの作成日時設定<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.SetLastWriteTime
'>* FileIO.SetLastAccessTime
'>
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
'>---
'>#### SetLastWriteTime(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.SetLastWriteTime(stFilePath, dtCreateTime)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|stFilePath|必須|String|更新日付を設定するファイルを指定|
'>|dtCreateTime|必須|更新日付を設定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル／フォルダの更新日時設定<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.SetCreationTime
'>* FileIO.SetLastAccessTime
'>
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
'>---
'>#### SetLastAccessTime(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.SetLastAccessTime(stFilePath, dtCreateTime)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|stFilePath|必須|String|アクセス日付を設定するファイルを指定|
'>|dtCreateTime|必須|アクセス日付を設定|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>ファイル／フォルダのアクセス日時設定<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* FileIO.SetCreationTime
'>* FileIO.SetLastWriteTime
'>
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
'>---
'>#### CreateFolderImage(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.CreateFolderImage(source, destination)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|source|必須|String|コピー元|
'>|destination|必須|コピー先|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>フォルダイメージコピー<br>
'>元フォルダと同じ構成のフォルダを作成する。<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
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
'>---
'>#### MoveFolder(FileIO)
'>
'>**Syntax**
'>
'>```
'>FileIO.MoveFolder(source, destination)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>|source|必須|String|移動元フォルダ|
'>|destination|必須|移動先フォルダ|
'>
'>**Return Value**
'>
'>* None
'>
'>**Remarks**
'>
'>フォルダの移動<br>
'>指定フォルダを移動する<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* None
'>
Sub MoveFolder(ByVal Source As String, ByVal destination As String)

    With FSO
        .MoveFolder Source, destination
    End With

End Sub

'アプリケーション名取得
Private Property Get Title() As String
    Title = ThisWorkbook.BuiltinDocumentProperties("Title").Value
End Property

