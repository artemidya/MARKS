Private Sub UserForm_Terminate()
    On Error Resume Next
    If CalculateSanP.SquareFormVisible = False Then
       CalculateSanP.Break
    Else
        If SquareForm.ComboBox1.ListIndex = 2 Then CalculateSanP.ErrHide
    End If
    CalculateSanP.PerimetrFormVisible = False
    
    If UserData.PerimetrFormData.Price <> ValP(TextBox1.text) * 100 _
    Or UserData.PerimetrFormData.Discount <> ValP(TextBox2.text) * 100 _
    Or UserData.PerimetrFormData.WindowPositionX <> CurWindowPositionX _
    Or UserData.PerimetrFormData.WindowPositionY <> CurWindowPositionY Then
        UserData.PerimetrFormData.Price = ValP(TextBox1.text) * 100
        UserData.PerimetrFormData.Discount = ValP(TextBox2.text) * 100
        UserData.PerimetrFormData.WindowPositionX = CurWindowPositionX
        UserData.PerimetrFormData.WindowPositionY = CurWindowPositionY
        UserData.Save
    End If
End Sub
