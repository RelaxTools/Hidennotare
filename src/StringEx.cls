VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEx"
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
'>### StringEx ƒNƒ‰ƒX
'>
'>**Remarks**
'>
'>Šg’£StringƒNƒ‰ƒX<br>
'>
'>**Dependent**
'>
'>* StringBuilder
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|V‹Kì¬|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit

Implements INewInstance

Private mBuf As String
Public Function NewInstance(Optional ByVal s As String = "") As StringEx
    Set NewInstance = New StringEx
    NewInstance.Value = s
End Function
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    'Attribute Item.VB_UserMemId = 0
    Value = mBuf
End Property
Public Property Let Value(v As String)
    mBuf = v
End Property
Public Function Length() As Long
    Length = StringUtils.LenEx(mBuf)
End Function
Public Function Trim() As StringEx
     Set Trim = Me.NewInstance(StringUtils.TrimEx(mBuf))
End Function
Public Function StartsWith(ByVal strWith As String) As Boolean
    StartsWith = StringUtils.StartsWith(mBuf, strWith)
End Function
Public Function EndsWith(ByVal strWith As String) As Boolean
    EndsWith = StringUtils.EndsWith(mBuf, strWith)
End Function
Public Function SubString(ByVal starts As Long, Optional ByVal ends As Long = -1) As StringEx
    If ends = -1 Then
        Set SubString = Me.NewInstance(Mid$(mBuf, starts + 1))
    Else
        Set SubString = Me.NewInstance(Mid$(mBuf, starts + 1, ends - starts))
    End If
End Function
Public Function Escape() As StringEx
    Set Escape = Me.NewInstance(StringUtils.Escape(mBuf))
End Function
Public Function Unescape() As StringEx
    Set Unescape = Me.NewInstance(StringUtils.Unescape(mBuf))
End Function
Public Function PlaceHolder(ParamArray p() As Variant) As StringEx
    
    '–|–ó
    Dim strMsg As String
    
    strMsg = StringUtils.Unescape(mBuf)
    
    If UBound(p) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(p)
            strMsg = VBA.Replace(strMsg, "{" & CStr(i) & "}", p(i))
        Next
    End If
    
    Set PlaceHolder = Me.NewInstance(strMsg)

End Function
Function FromUTF8(ByRef bytBuf() As Byte)
    mBuf = StringUtils.FromUTF8(bytBuf)
End Function
Function ToUTF8() As Byte()
    ToUTF8 = StringUtils.ToUTF8(mBuf)
End Function
Public Function ToHiragana(Optional ByVal flag As Boolean = False) As StringEx
    Set ToHiragana = Me.NewInstance(StringUtils.ToHiragana(mBuf, flag))
End Function
Public Function ToKatakana(Optional ByVal flag As Boolean = False) As StringEx
    Set ToKatakana = Me.NewInstance(StringUtils.ToKatakana(mBuf, flag))
End Function
Private Function INewInstance_NewInstance() As Object
    Set INewInstance_NewInstance = Me.NewInstance()
End Function
Public Function IndexOf(ByVal ch As String) As Long
    IndexOf = InStr(1, mBuf, ch, vbBinaryCompare) - 1
End Function
Public Function Matches(ByVal regex As String) As Boolean
    Matches = RegExp.Test(mBuf, regex)
End Function
Public Function Replace(ByVal exp As String, ByVal rep As String, Optional ByVal ignorecase As Boolean = False) As StringEx
    Set Replace = Me.NewInstance(RegExp.Replace(mBuf, exp, rep, ignorecase))
End Function
Public Function ToUpperCase() As StringEx
    Set ToUpperCase = Me.NewInstance(UCase(mBuf))
End Function
Public Function ToLowerCase() As StringEx
    Set ToLowerCase = Me.NewInstance(LCase(mBuf))
End Function
Public Function CharAt(ByVal pos As Long) As StringEx
    Set CharAt = Me.NewInstance(Mid$(mBuf, pos + 1, 1))
End Function
Public Function GetBytes() As Byte()
    GetBytes = mBuf
End Function

