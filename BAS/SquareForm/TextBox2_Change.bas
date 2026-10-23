Private Sub TextBox2_Change()
    If IsEditingMyFields = False Then
        UserFormRefreshData
        If CalculateSanP.PerimetrFormVisible = True Then
            PerimetrForm.IsEditingMyFields = True
            PerimetrForm.TextBox2 = TextBox2
            PerimetrForm.IsEditingMyFields = False
            PerimetrForm.UserFormRefreshData
        End If
    End If
End Sub
