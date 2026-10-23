Private Sub UserForm_Activate()
    If FirstRun = True Then
        If CalculateSanP.PerimetrFormVisible = False Then
            CalculateSanP.Start
        Else
            Me.Repaint
            If ComboBox1.ListIndex = 1 Then CalculateSanP.ShowRectOnScreenCurves
            If ComboBox1.ListIndex = 2 Then CalculateSanP.SelectionRectOnScreenCurve.Show
        End If
    End If
    FirstRun = False
End Sub
