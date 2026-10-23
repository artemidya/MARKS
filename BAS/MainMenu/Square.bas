Public Sub Square()
    On Error GoTo 10
    If CalculateSanP.SquareFormVisible Then
        Unload SquareForm
    Else
        SquareForm.Show 0
    End If
10
End Sub
