Public Function MovePU(x As Long, y As Long) As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MovePU = "PU" & CStr(x) & "," & CStr(y)
        Case DMPL
            MovePU = "U" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function
