Private Sub CommandButton1_Click()
    Dim txt As String
    Dim bytes() As Byte
    
    txt = "IN;PA;VS5;PU0,0;PU0,400;PU400,400;PU400,0;PU0,0;"
    bytes = PlotterLanguage.StrToAscii(txt)
    PlotterForm.OutPuts.Out bytes, PlotterForm.OutPuts.x, PlotterForm.OutPuts.y, True
End Sub
