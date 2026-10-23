Private Sub CommandButton2_Click()
    ListenKeyboard = False
    OutPuts.Configure KillBirdsAndSpaces(Combo1.text), FindWindow(vbNullString, Me.Caption)
End Sub
