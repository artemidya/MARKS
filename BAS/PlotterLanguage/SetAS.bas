Public Function SetAS(acceleration As Long) As String
    If acceleration = 0 Then Exit Function
    'as = 1-4
    ' 1 -> 1,25 m/s^2
    ' 2 -> 2,5 m/s^2
    ' 3 -> 5 m/s^2
    ' 4 -> 10 m/s^2
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetAS = "AS" & CStr(acceleration)
        Case DMPL
            'Exit (no acceleration command in DMPL)
    End Select
End Function
