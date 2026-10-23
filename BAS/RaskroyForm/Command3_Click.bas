Private Sub Command3_Click()
    SelectedResult = SelectedResult - 1
    If SelectedResult = 1 Then Command3.Enabled = False
    If CountResults > 1 Then Command6.Enabled = True
    If SelectedResult = CountResults Then Command6.Enabled = False
    If Not Plotf Then Command4.Enabled = True
    Label1.Caption = SelectedResult & " " & localization.ActiveLng.Item("RaskroyForm.From") & " " & CountResults & "   (" & Int(Work_Result(SelectedResult).maximumY * 25.4) & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    Dim wrtemp As Work_Result
    wrtemp = Work_Result(SelectedResult)
    DrawWarkResult wrtemp
End Sub
