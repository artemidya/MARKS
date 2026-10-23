Public Sub SetTexts()
    PerimetrForm.Caption = localization.ActiveLng.Item("PerimetrForm.Caption")
    Label1.Caption = localization.ActiveLng.Item("PerimetrForm.Label1.Caption")
    Label8.Caption = localization.ActiveLng.Item("PerimetrForm.Label8.Caption")
    Label9.Caption = localization.ActiveLng.Item("PerimetrForm.Label9.Caption")
    
    Label3.Caption = localization.ActiveLng.Item("PerimetrForm.Label3.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit) & "/" & localization.ActiveLng.Item("PerimetrForm.m")
    Label5.Caption = localization.ActiveLng.Item("PerimetrForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("PerimetrForm.Label7.Caption") & ", " & Trim(UserData.PerimetrFormData.MoneyUnit)
End Sub
