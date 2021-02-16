Attribute VB_Name = "Module1"
'一年分のシートを印刷するマクロ
Sub OneYearsWorthOfPrinting()
    Dim tsukihazime As Date '月初め
    Dim nen As Integer '入力された西暦
    Dim i As Integer 'カウンタ
    Dim insatsuPreview As Boolean '印刷プレビュー
    
    
    nen = Application.InputBox("年度を西暦で入力してください" & vbNewLine & "例：2021", Type:=1)
              
    If nen = 0 Then 'キャンセルボタンが押された
        MsgBox "キャンセルしました"
        End '終了
    End If
              
    tsukihazime = DateValue(nen & "/04/01") '年度初めの指定
    a = Application.Dialogs(xlDialogPrinterSetup).Show 'プリンターの選択
    
    'プリンター選択の例外処理
    If Not a Then
        MsgBox "中断します"
        End '終了
    End If
    
    '印刷プレビューの設定
    Dim rslt As VbMsgBoxResult
    rslt = MsgBox("印刷プレビューを表示しますか？" & vbNewLine & "（「はい」を推奨）", Buttons:=vbYesNo)
    If rslt = vbYes Then
        insatsuPreview = True
    Else
        insatsuPreview = False
    End If

    For i = 1 To 12
        Range("A1").Value = tsukihazime 'セルに入力（????/??/01）
        On Error GoTo printError 'エラーが発生したら飛ぶ
        ActiveSheet.PrintOut Preview:=insatsuPreview '印刷
        tsukihazime = DateAdd("m", 1, tsukihazime) 'インクリメント
    Next
    
    Range("A1").Value = tsukihazime '次年度をセット
    
Exit Sub

printError:
    MsgBox "印刷でエラー"
    
End Sub
