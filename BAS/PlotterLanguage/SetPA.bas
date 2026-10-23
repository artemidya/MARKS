Public Function SetPA() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetPA = "PA"
        Case DMPL
            SetPA = "A"
    End Select
End Function
