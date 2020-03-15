VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWait 
   Caption         =   "UserForm1"
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9540.001
   OleObjectBlob   =   "frmWait.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit
Implements IUsing

Private mlngMax As Long
Private m_Cancel As Boolean

Private Sub UserForm_Initialize()
    m_Cancel = False
End Sub
'--------------------------------------------------------------
' タイトルバー設定
'--------------------------------------------------------------
Public Property Let TitleBar(ByVal v As String)
    Me.Caption = v
End Property
'--------------------------------------------------------------
' メッセージ設定
'--------------------------------------------------------------
Public Property Let Message(ByVal v As String)
    lblMessage.Caption = v
End Property
'--------------------------------------------------------------
' キャンセルプロパティ
'--------------------------------------------------------------
Public Property Get IsCancel() As Boolean
    IsCancel = m_Cancel
End Property
'--------------------------------------------------------------
'キャンセルボタン
'--------------------------------------------------------------
Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    DoEvents
    m_Cancel = True
End Sub
'--------------------------------------------------------------
'閉じるボタン無効
'--------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Select Case CloseMode
        Case 0
            Cancel = True
    End Select

End Sub
'--------------------------------------------------------------
'  ガイダンスメッセージ表示
'--------------------------------------------------------------
Public Sub DispGuidance(ByVal strValue As String)

    lblBar.Caption = strValue
    lblStatus.Caption = strValue
    DoEvents

End Sub
'--------------------------------------------------------------
'  進捗状況開始
'--------------------------------------------------------------
Public Sub StartGauge(ByVal lngValue As Long)
    
    mlngMax = lngValue
    lblBar.Width = 0
    lblBar.visible = True

End Sub
'--------------------------------------------------------------
'  進捗状況描画
'--------------------------------------------------------------
Public Sub DisplayGauge(ByVal lngValue As Long)

    Dim dblValue As Double
    Dim strMessage As String
    
    If lngValue > mlngMax Then
        lngValue = mlngMax
    End If
    dblValue = (CDbl(lngValue) / mlngMax)
    lblBar.Width = lblStatus.Width * dblValue
    
    strMessage = Space$(Fix(lblStatus.Width * 0.16)) & Format$(Fix(dblValue * 100), "0") & "%"
    lblBar.Caption = strMessage
    lblStatus.Caption = strMessage
    DoEvents
    
End Sub
'--------------------------------------------------------------
' IUsing I/F Begin
'--------------------------------------------------------------
Private Sub IUsing_Begin()
    m_Cancel = False
    Me.Show
End Sub
'--------------------------------------------------------------
' IUsing I/F Finish
'--------------------------------------------------------------
Private Sub IUsing_Finish()
    Unload Me
End Sub

