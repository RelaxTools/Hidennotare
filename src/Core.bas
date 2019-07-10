Attribute VB_Name = "Core"
'-----------------------------------------------------------------------------------------------------
'
' [hidennotare] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/hidennotare
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
' コンストラクタ生成
'-----------------------------------------------------------------------------------------------------
Option Explicit

'Callback用
Private mCallback As IDictionary

'-----------------------------------------------------------------------------------------------------
' コンストラクタ生成
'-----------------------------------------------------------------------------------------------------
Public Function Constructor(ByRef obj As Object, ParamArray Args() As Variant) As Object

    Dim c As IConstructor
    Dim v As Variant
    
    'コレクションのコンストラクタ
    If TypeOf obj Is Collection Then
        For Each v In Args
            obj.Add v
        Next
        Set Constructor = obj
    
    'その他クラスのコンストラクタ
    Else
        '引数をCollectionに詰め替える
        Dim col As Collection
        Set col = New Collection
        
        For Each v In Args
            'FormのMe指定の場合Controlsが入ってしまう対策
            If TypeName(v) = "Controls" Then
                col.Add v(1).Parent
            Else
                col.Add v
            End If
        Next
        
        'IConstructor Interfaceを呼び出す。
        Set c = obj
        Set Constructor = c.Instancing(col)
    End If
    
    'オブジェクトが返却されなかった場合エラー
    If Constructor Is Nothing Then
        Message.Throw 1, "ClassHelper", "Constructor", "Argument Error"
    End If

End Function
'-----------------------------------------------------------------------------------------------------
' VBA 個人的汎用処理 https://qiita.com/nukie_53/items/bde16afd9a6ca789949d
' @nukie_53
' Set/Letを隠蔽するプロパティ
'-----------------------------------------------------------------------------------------------------
Public Property Let SetVar(outVariable As Variant, inExpression As Variant)
    
    Select Case True
        Case VBA.IsObject(inExpression), VBA.VarType(inExpression) = vbDataObject
            
            Set outVariable = inExpression
        
        Case Else
            
            Let outVariable = inExpression
    
    End Select

End Property
'-----------------------------------------------------------------------------------------------------
' ソースのエクスポート
'-----------------------------------------------------------------------------------------------------
Sub Export()

    Dim strFile As String
    Dim strExt As String
    Dim obj As Object
    Dim strTo As String
    
    strFile = FileIO.BuildPath(FileIO.GetParentFolderName(ActiveWorkbook.FullName), "src")
    FileIO.CreateFolder strFile
    
    For Each obj In ActiveWorkbook.VBProject.VBComponents
    
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
    
    MsgBox "ソースを保存しました。", vbInformation, "Export"
    
End Sub
'---------------------------------------------------------------------------------------------------
'　Callbackの際のInstallメソッド
'---------------------------------------------------------------------------------------------------
Public Function InstallCallback(MH As Callback) As String

    Dim Key As String

    If mCallback Is Nothing Then
        Set mCallback = New Dictionary
    End If
    
    Key = CStr(ObjPtr(MH))
    
    mCallback.Add Key, MH
    
    InstallCallback = Key
    
End Function
'---------------------------------------------------------------------------------------------------
'　Callbackの際のUnInstallメソッド
'---------------------------------------------------------------------------------------------------
Public Sub UninstallCallback(ByVal Key As String)

    If mCallback.ContainsKey(Key) Then
        mCallback.Remove Key
    End If
    
End Sub
'---------------------------------------------------------------------------------------------------
'　Callbackの際に呼び出されるメソッド
'---------------------------------------------------------------------------------------------------
Public Function OnActionCallback(ByVal Key As String, ByVal lngEvent As Long, ByVal opt As String)

    Dim MH As Callback
    
    If mCallback.ContainsKey(Key) Then
        
        Set MH = mCallback(Key)
        Call MH.OnActionCallback(lngEvent, opt)
        
    End If

End Function

