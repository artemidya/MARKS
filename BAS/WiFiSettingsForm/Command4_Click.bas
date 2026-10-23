Private Sub Command4_Click()
    'Save
    On Error GoTo 10
    PlotterForm.RegData1.IP(PlotterForm.TabStrip1.SelectedItem.index + 1) = TextBox1.text & "." & TextBox2.text & "." & TextBox3.text & "." & TextBox4.text
    PlotterForm.RegData1.Port(PlotterForm.TabStrip1.SelectedItem.index + 1) = CInt(TextBox5.text)
    PlotterForm.OutPuts.WiFi_Refresh
10
    Unload Me
End Sub
