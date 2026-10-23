Public Sub Perimetr()
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Then
        Unload PerimetrForm
    Else
        PerimetrForm.Show 0
    End If
10
End Sub
