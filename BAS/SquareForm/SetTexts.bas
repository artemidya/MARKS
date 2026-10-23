Public Sub SetTexts()
    SquareForm.Caption = localization.ActiveLng.Item("SquareForm.Caption")
    Label3.Caption = localization.ActiveLng.Item("SquareForm.Label3.Caption") & ChrW(178) & ", " & Trim(UserData.PerimetrFormData.MoneyUnit) & "/" & localization.ActiveLng.Item("SquareForm.m") & ChrW(178)
    Label5.Caption = localization.ActiveLng.Item("SquareForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("SquareForm.Label7.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit)
    ComboBox1.List(0) = localization.ActiveLng.Item("SquareForm.AreaShapes") & ChrW(178)
    ComboBox1.List(1) = localization.ActiveLng.Item("SquareForm.AreaOfRectangularBars") & ChrW(178)
    ComboBox1.List(2) = localization.ActiveLng.Item("SquareForm.AreaOfSelection") & ChrW(178)
End Sub
