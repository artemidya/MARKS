Public Function SetVS(velocity As Long) As String
    velocity = velocity \ 10
    If velocity = 0 Then Exit Function
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetVS = "VS" & CStr(velocity)
        Case DMPL
            SetVS = "V" & CStr(velocity)
    End Select
End Function
