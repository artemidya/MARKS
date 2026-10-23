Private Sub UserForm_Terminate()
    On Error Resume Next
    If CalculateSanP.PerimetrFormVisible = False Then
        CalculateSanP.Break
    End If
    CalculateSanP.SquareFormVisible = False
    CalculateSanP.HideRectOnScreenCurves
    CalculateSanP.SelectionRectOnScreenCurve.Hide
    
    If UserData.SquareFormData.Price <> ValP(TextBox1.text) * 100 _
    Or UserData.PerimetrFormData.Discount <> ValP(TextBox2.text) * 100 _
    Or UserData.SquareFormData.WindowPositionX <> CurWindowPositionX _
    Or UserData.SquareFormData.WindowPositionY <> CurWindowPositionY _
    Or UserData.SquareFormData.RectArea <> ComboBox1.ListIndex Then
        UserData.SquareFormData.Price = ValP(TextBox1.text) * 100
        UserData.PerimetrFormData.Discount = ValP(TextBox2.text) * 100
        UserData.SquareFormData.WindowPositionX = CurWindowPositionX
        UserData.SquareFormData.WindowPositionY = CurWindowPositionY
        UserData.SquareFormData.RectArea = ComboBox1.ListIndex
        UserData.Save
    End If
End Sub
