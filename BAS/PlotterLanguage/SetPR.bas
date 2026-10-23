Public Function SetPR() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            SetPR = "PR"
        Case DMPL
            SetPR = "R"
    End Select
End Function
