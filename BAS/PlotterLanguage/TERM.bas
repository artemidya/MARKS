Public Function TERM() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            TERM = ""
        Case DMPL
            TERM = "U@"
    End Select
End Function
