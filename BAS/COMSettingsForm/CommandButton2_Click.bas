Private Sub CommandButton2_Click()
    Dim txt As String
    Dim bytes() As Byte
    
    txt = ";:HL0,EC1,UP1,AV5,U0,0,U0,400,U400,400,U400,0,U0,0,"
    bytes = PlotterLanguage.StrToAscii(txt)
    PlotterForm.OutPuts.Out bytes, PlotterForm.OutPuts.x, PlotterForm.OutPuts.y, True
End Sub
