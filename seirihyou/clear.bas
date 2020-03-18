Attribute VB_Name = "Module2"
Sub レコードの全削除()
'
' レコードの全削除 Macro
'

'
    Dim ans As Integer
    ans = MsgBox("全てのレコードを削除しますか？ ---この操作は元に戻せません。---", vbYesNo + vbExclamation, "警告：重大な操作")
    
    If ans = vbYes Then
        ans = MsgBox("本当に削除しますか？", vbYesNo, "最終確認")
        
        If ans = vbYes Then
            
            ' 上書き保存
            ActiveWorkbook.Save
            
            ' バックアップファイルの作成
            Dim originalFileName As String
            originalFileName = ActiveWorkbook.FullName
            Dim backupFileName As String
            backupFileName = ActiveWorkbook.FullName & ".backup"
            ActiveWorkbook.SaveAs Filename:=backupFileName
            
            ' オリジナルファイルの展開
            backupFileName = ActiveWorkbook.Name
            Workbooks.Open (originalFileName)
            originalFileName = Dir(originalFileName)
            Workbooks(originalFileName).Activate
            
            ' レコードの削除
            Range("B4:N104").Select
            Selection.ClearContents
            ans = MsgBox("レコードを削除しました。", vbOKOnly + vbInformation, "初期化の完了")
            Cells(4, 2).Select
            
            ' バックアップファイルのクローズ
            Workbooks(backupFileName).Close
        
        Else
            ans = MsgBox("操作はキャンセルされました。", vbOKOnly, "初期化の中止")
        
        End If
    
    Else
        ans = MsgBox("操作はキャンセルされました。", vbOKOnly, "初期化の中止")
    
    End If
   
End Sub


