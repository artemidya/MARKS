Private Sub UserForm_Initialize()
    Skip = False
    MbClosing = False
    Break = False
    ProgressForm.PBar.Caption = ""
    ProgressForm.PBarStrip.Width = 0
    k = ProgressForm.PBar.Width
    CommandButton1.Caption = localization.ActiveLng.Item("ProgressForm.Skip") & " >>"
End Sub
