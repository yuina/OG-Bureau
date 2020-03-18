Attribute VB_Name = "Module1"
Sub 工事一覧を表示()
Attribute 工事一覧を表示.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 工事一覧を表示 Macro
'

'
    ActiveSheet.Range("$B$4:$B104").AutoFilter Field:=2, Criteria1:="=*工事*", _
        Operator:=xlAnd
End Sub

Sub 業務一覧を表示()
'
' 業務一覧を表示 Macro
'

'
    ActiveSheet.Range("$B$4:$B104").AutoFilter Field:=2, Criteria1:="=*業務*", _
        Operator:=xlAnd
End Sub

Sub フィルターのクリア()
Attribute フィルターのクリア.VB_ProcData.VB_Invoke_Func = " \n14"
'
' フィルターのクリア Macro
'

'
    ActiveSheet.Range("$B$4:$B104").AutoFilter Field:=2
End Sub

