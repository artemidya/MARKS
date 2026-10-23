Public Function MovePD(x As Long, y As Long) As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            MovePD = "PD" & CStr(x) & "," & CStr(y)
        Case DMPL
            MovePD = "D" & CStr(x) & "," & CStr(y) & ","
    End Select
End Function
