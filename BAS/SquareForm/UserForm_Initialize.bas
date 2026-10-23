Private Sub UserForm_Initialize()
    UserData.Load
    FirstRun = True
    CalculateSanP.SquareFormVisible = True

    left = UserData.SquareFormData.WindowPositionX
    CurWindowPositionX = left
    top = UserData.SquareFormData.WindowPositionY
    CurWindowPositionY = top
    
    IsEditingMyFields = True
    
    ComboBox1.AddItem
    ComboBox1.AddItem
    ComboBox1.AddItem
    SetTexts


    If UserData.SquareFormData.RectArea = 1 Then
        ComboBox1.ListIndex = 1
        Image3.Picture = Image2.Picture
    ElseIf UserData.SquareFormData.RectArea = 2 Then
        ComboBox1.ListIndex = 2
        Image3.Picture = Image4.Picture
    Else
        ComboBox1.ListIndex = 0
        Image3.Picture = Image1.Picture
    End If
    

    
    TextBox1 = Format(UserData.SquareFormData.Price / 100, "0.00")
    If CalculateSanP.PerimetrFormVisible = True Then
        TextBox2 = PerimetrForm.TextBox2
    Else
        TextBox2 = Format(UserData.PerimetrFormData.Discount / 100)
    End If
    IsEditingMyFields = False
    

    
    TextBox1.SelStart = 0
    TextBox1.SelLength = 10
    UserFormRefreshData
End Sub
