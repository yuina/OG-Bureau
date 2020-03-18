Attribute VB_Name = "Module1"
Sub 抽出データの保存()
'
' 抽出データの保存 Macro
'

'
    Dim rslt As VbMsgBoxResult
    rslt = MsgBox("KKMSから出力したCSVファイルを開いていますか？", Buttons:=vbYesNo)
    If rslt = vbYes Then
        On Error GoTo myError
        Windows("select.csv").Activate
        On Error GoTo 0
        Sheets("select").Select
        Sheets("select").Move After:=Workbooks(ThisWorkbook.Name).Sheets(5)
        リストとの一致判定
        MsgBox "処理が完了しました。" & vbCrLf & "シートの名前を変更後、保存してください。", vbExclamation
    Else
        MsgBox "KKMSで出力した「select.csv」を開いてから実行してください。"
    End If
    Exit Sub ' 正常時処理の終了
    
myError:
    MsgBox "「select.csv」を開いていません。" & vbCrLf & "KKMSから出力したファイルを開いているにもかかわらずこのエラーが発生する場合は、すべてのエクセルを閉じた後に、もう一度やり直してください。", vbCritical
    
End Sub

Sub リストとの一致判定()
'
' リストとの一致判定 Macro
'

'
    ' 列の挿入
    Dim addSheet
    addSheet = ActiveSheet.Name
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    ' 判定数式の入力
    Range("A2").Select
    ActiveCell.Formula = "=IF(COUNTIF(リスト!$B:$B,$C2),""あり"",""なし"")"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A300"), Type:=xlFillDefault
    Range("A2:A300").Select
    
    ' 色づけ
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="なし", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Cells(1, 1).Value = "リストとの重複"
    Sheets(addSheet).Activate
End Sub

Sub 一致判定の準備()
'
' 一致判定の準備 Macro
'

'
    ' 列の挿入
    Sheets("select").Activate
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    
    ' 判定数式の入力
    Range("A2").Select
    ActiveCell.Formula = "=IF(COUNTIF(リスト!$B:$B,$C2),""あり"",""なし"")"
    Range("A2").Select
    Selection.AutoFill Destination:=Range("A2:A300"), Type:=xlFillDefault
    Range("A2:A300").Select
    
    ' 色づけ
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="なし", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Cells(1, 1).Value = "リストとの重複"
    Sheets("select").Activate
    MsgBox "「select」に対して「リスト」との重複判定を行いました。"
End Sub

Function LastSaveTime() As Variant
Application.Volatile
LastSaveTime = ThisWorkbook.BuiltinDocumentProperties("Last save time").Value
End Function

