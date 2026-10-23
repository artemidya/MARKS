Private Sub ComboBox1_Change()
    If IsEditingMyFields = False Then
        If ComboBox1.ListIndex = 0 Then
            Image3.Picture = Image1.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.HideRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Hide
        ElseIf ComboBox1.ListIndex = 1 Then
            Image3.Picture = Image2.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.ShowRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Hide
        Else
            Image3.Picture = Image4.Picture
            UserFormRefreshData
            Me.Repaint
            CalculateSanP.HideRectOnScreenCurves
            CalculateSanP.SelectionRectOnScreenCurve.Show
        End If
    End If
End Sub
