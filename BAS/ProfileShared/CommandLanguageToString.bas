Public Function CommandLanguageToString(cl As CommandLanguages) As String
    Select Case cl
        Case HPGL
            CommandLanguageToString = "HPGL"
        Case DMPL
            CommandLanguageToString = "DMPL"
        Case Else
            CommandLanguageToString = "HPGL"
    End Select
End Function
