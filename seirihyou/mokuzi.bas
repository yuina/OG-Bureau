Attribute VB_Name = "Module3"
Sub 目次の作成()
Attribute 目次の作成.VB_Description = "目次作成\nそれぞれのシートでテキストフィルターの適用"
Attribute 目次の作成.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 目次の作成 Macro
' 目次作成 それぞれのシートでテキストフィルターの適用
'

'
    Sheets("工事目次").Range("$A$1:$F$101").AutoFilter Field:=3, Criteria1:="=*工事*", _
        Operator:=xlOr
    Sheets("業務目次").Range("$A$1:$F$101").AutoFilter Field:=3, Criteria1:="=*業務*", _
        Operator:=xlOr
    Dim ans As Integer
    ans = MsgBox("「工事」と「業務」の目次を作成しました。", vbOKOnly, "フィルターの完了")
End Sub
