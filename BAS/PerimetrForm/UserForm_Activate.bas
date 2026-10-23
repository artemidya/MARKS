Private Sub UserForm_Activate()
    If CalculateSanP.SquareFormVisible = False Then
       If FirstRun = True Then
            CalculateSanP.Start
       End If
    End If
    FirstRun = False
End Sub
