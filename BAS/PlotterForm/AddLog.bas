Public Sub AddLog(txt As String, Optional isInput As Boolean = False)
    On Error GoTo 10
    TextBox11.SetFocus
    txt = Replace(txt, vbNewLine, "")
    txt = Trim(txt)
    If txt = "" Then GoTo 10
    
    If isInput Then
        TextBox11.text = TextBox11.text & KillBirdsAndSpaces(Combo1.text) & " >> "
    Else
        TextBox11.text = TextBox11.text & KillBirdsAndSpaces(Combo1.text) & " << "
    End If
    TextBox11.text = TextBox11.text & txt & vbNewLine
10
End Sub
