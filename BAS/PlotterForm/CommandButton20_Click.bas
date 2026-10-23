Private Sub CommandButton20_Click()
    Dim txt As String
    txt = UCase(TextBox12.text)
    Dim bytes() As Byte
    bytes = StrToAscii(txt)
    OutPuts.Out bytes, OutPuts.x, OutPuts.y
    SendKeys "{PGDN}"
    TextBox12.text = ""
    TextBox12.SetFocus
    
End Sub
