VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTest 
   Caption         =   "Hdennotare Test"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310.001
   OleObjectBlob   =   "frmTest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim FM As IFormManager

Private Sub IList_Add(obj As Variant)

End Sub

Private Sub IList_Clear()

End Sub

Private Sub UserForm_Initialize()

    lblBack.Tag = "m"
    lblGauge.Tag = "g"
    cmdOk.Tag = "c"

    Set FM = FormManager.NewInstance(Me)

    Dim i As Long
    Dim strBuf As String

    With lvTest
        
        .View = lvwReport           ''表示
        .LabelEdit = lvwManual      ''ラベルの編集
        .HideSelection = False      ''選択の自動解除
        .AllowColumnReorder = False  ''列幅の変更を許可
        .FullRowSelect = True       ''行全体を選択
        .Gridlines = True           ''グリッド線
 
        .ColumnHeaders.Add , "_Name", "メソッド", .Width - 16
  
    End With
  
    With ThisWorkbook.VBProject.VBComponents("Test").CodeModule
            
        For i = 1 To .CountOfLines
            
            '指定位置から１行取得
            strBuf = .Lines(i, 1)
            
            If RegExp.Test(strBuf, "^Sub Test.*\)$") Then
            
                With lvTest.ListItems.Add
                    .Text = Replace(Mid$(strBuf, 5), "()", "")
                End With
            
            End If
    
        Next
    End With

    FM.DispGuidance "テストを読み込みました。"

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '実行中か？
    If FM.IsRunning Then
        'フォームが閉じられないようにする
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Set FM = Nothing
End Sub
Private Sub cmdOk_Click()

    '実行中か？実行中は中断ボタンになるので、キャンセル動作。
    If FM.IsRunning Then
        'キャンセルを実行する
        FM.doCancel
        Exit Sub
    End If

    If Message.Question("実行します。よろしいですか？") Then
        Exit Sub
    End If

    Dim lngMax As Long
    Dim i As Long
    
    With Using.NewInstance(FM, New OneTimeSpeedBooster)
        
        lngMax = lvTest.ListItems.Count
    
        'ゲージの最大を設定
        FM.StartGauge lngMax
        
        For i = 1 To lngMax
        
            'キャンセルか？
            If FM.IsCancel Then
                '処理を中断
                Exit For
            End If
        
            '実行
            lvTest.ListItems(i).Selected = True
            lvTest.ListItems(i).EnsureVisible
            
            Application.Run "Test." & lvTest.ListItems(i).Text
            
            Process.Sleep 50
            DoEvents
                    
            'ゲージの現在値を設定
            FM.DisplayGauge i
        Next
    
    End With
    
    'キャンセルか？
    If FM.IsCancel Then
        FM.DispGuidance "テストは中断されました。"
    Else
        FM.DispGuidance "テストが完了しました。"
    End If

End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

