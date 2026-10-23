Public Function Init() As String
    Select Case Profiles(PlotterForm.TabStrip1.SelectedItem.index).CommandLanguage
        Case HPGL
            Init = "IN;"
        Case DMPL
            Init = ";:HL0,EC1,UP1,"
    End Select
End Function
