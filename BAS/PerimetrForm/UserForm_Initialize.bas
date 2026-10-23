Private Sub UserForm_Initialize()
    UserData.Load
    FirstRun = True
    CalculateSanP.PerimetrFormVisible = True
    
    left = UserData.PerimetrFormData.WindowPositionX
    CurWindowPositionX = left
    top = UserData.PerimetrFormData.WindowPositionY
    CurWindowPositionY = top
    SetTexts
    
    IsEditingMyFields = True
    TextBox1 = Format(UserData.PerimetrFormData.Price / 100, "0.00")
    If CalculateSanP.SquareFormVisible = True Then
        TextBox2 = SquareForm.TextBox2
    Else
        TextBox2 = Format(UserData.PerimetrFormData.Discount / 100)
    End If
    IsEditingMyFields = False
    
    TextBox1.SelStart = 0
    TextBox1.SelLength = 10
    CalculateSanP.ErrShow
    UserFormRefreshData
End Sub
