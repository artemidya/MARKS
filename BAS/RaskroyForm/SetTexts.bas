Private Sub SetTexts()
    Frame2.Caption = localization.ActiveLng.Item("RaskroyForm.Frame2.Caption")
    CommandNew.Caption = localization.ActiveLng.Item("RaskroyForm.CommandNew.Caption")
    CommandSave.Caption = localization.ActiveLng.Item("RaskroyForm.CommandSave.Caption")
    CommandDelete.Caption = localization.ActiveLng.Item("RaskroyForm.CommandDelete.Caption")
    MultiPage1.Pages("Page1").Caption = localization.ActiveLng.Item("RaskroyForm.MultiPage1.Pages.Page1.Caption")
    MultiPage1.Pages("Page2").Caption = localization.ActiveLng.Item("RaskroyForm.MultiPage1.Pages.Page2.Caption")
    Frame1.Caption = localization.ActiveLng.Item("RaskroyForm.Frame1.Caption")
    Label3.Caption = localization.ActiveLng.Item("RaskroyForm.Label3.Caption")
    Label4.Caption = localization.ActiveLng.Item("RaskroyForm.Label4.Caption")
    Check2.Caption = localization.ActiveLng.Item("RaskroyForm.Check2.Caption")
    Frame3.Caption = localization.ActiveLng.Item("RaskroyForm.Frame3.Caption")
    Label6.Caption = localization.ActiveLng.Item("RaskroyForm.Label6.Caption")
    Label5.Caption = localization.ActiveLng.Item("RaskroyForm.Label5.Caption")
    Label7.Caption = localization.ActiveLng.Item("RaskroyForm.Label7.Caption")
    Check1.Caption = localization.ActiveLng.Item("RaskroyForm.Check1.Caption")
    Label8.Caption = localization.ActiveLng.Item("RaskroyForm.Label8.Caption")
    Label9.Caption = localization.ActiveLng.Item("RaskroyForm.Label9.Caption")
    
    Check4.Caption = localization.ActiveLng.Item("RaskroyForm.Check4.Caption")
    Check3.Caption = localization.ActiveLng.Item("RaskroyForm.Check3.Caption")
    Check5.Caption = localization.ActiveLng.Item("RaskroyForm.Check5.Caption")
    Check6.Caption = localization.ActiveLng.Item("RaskroyForm.Check6.Caption")
    Check7.Caption = localization.ActiveLng.Item("RaskroyForm.Check7.Caption")
    Check8.Caption = localization.ActiveLng.Item("RaskroyForm.Check8.Caption")
    Check9.Caption = localization.ActiveLng.Item("RaskroyForm.Check9.Caption")
    CheckBox1.Caption = localization.ActiveLng.Item("RaskroyForm.CheckBox1.Caption")
    dntChangeCheckBox1 = True
    CheckBox1.value = RegData1.NestingPlaceOnPage
    dntChangeCheckBox1 = False

    
    Command4.Caption = localization.ActiveLng.Item("RaskroyForm.Command4.Caption")
    Command5.Caption = localization.ActiveLng.Item("RaskroyForm.Command5.Caption")
    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Command1.Caption")
    
    Combo2.List(0) = localization.ActiveLng.Item("RaskroyForm.withoutTurning")
End Sub
