Private Sub EnableAllButtons()
    EnebleProfileAllElements
    Combo1.Enabled = True
    CommandNew.Enabled = True
    If NeedSave Then CommandSave.Enabled = True
    CommandDelete.Enabled = True
    If SelectedResult > 1 Then Command3.Enabled = True
    If SelectedResult < CountResults Then Command6.Enabled = True
    If CountResults > 0 And Plotf = False Then Command4.Enabled = True
    Frame2.Enabled = True
End Sub
