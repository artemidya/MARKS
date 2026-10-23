Private Sub SetAllTexts()
    If CalculateSanP.PerimetrFormVisible Then
        PerimetrForm.UserFormRefreshData
        PerimetrForm.SetTexts
    End If
    If CalculateSanP.SquareFormVisible Then
        SquareForm.UserFormRefreshData
        SquareForm.SetTexts
    End If
End Sub
