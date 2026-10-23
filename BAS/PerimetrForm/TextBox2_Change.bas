Private Sub TextBox2_Change()
    If IsEditingMyFields = False Then
        UserFormRefreshData
        If CalculateSanP.SquareFormVisible = True Then
            SquareForm.IsEditingMyFields = True
            SquareForm.TextBox2 = TextBox2
            SquareForm.IsEditingMyFields = False
            SquareForm.UserFormRefreshData
        End If
    End If
End Sub
