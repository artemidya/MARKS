Private Sub SelectLangInCombo(Lng As String)
    On Error GoTo 10
    Combo1.value = UCase(Lng)
    Exit Sub
10:
    Combo1.value = "EN"
End Sub
