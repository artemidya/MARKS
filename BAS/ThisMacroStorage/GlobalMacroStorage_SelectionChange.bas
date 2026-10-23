Private Sub GlobalMacroStorage_SelectionChange()
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Or CalculateSanP.SquareFormVisible Then
        CalculateSanP.Start
    End If
10
End Sub
