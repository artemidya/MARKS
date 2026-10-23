Private Sub RefreshCommand1()
    Command1.Caption = localization.ActiveLng.Item("RaskroyForm.Command1.Caption")
    Label10.Caption = ""
    Command1.BackColor = VBA.SystemColorConstants.vbButtonFace
    ProgressBar1.GreenColorOn = False
    ImproveOn = False
    EnableCommand1
    Label1.Caption = ".."
End Sub
