Public Function StringToCommandLanguage(text As String) As CommandLanguages
    text = UCase(text)
    Select Case text
        Case "HPGL"
            StringToCommandLanguage = HPGL
        Case "DMPL"
            StringToCommandLanguage = DMPL
        Case Else
            StringToCommandLanguage = HPGL
    End Select
End Function
