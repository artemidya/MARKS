Property Get ComSettingsStr() As String
    On Error GoTo 10
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Interface
            ComSettingsStr = Interfaces.ComSettingsStr(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            ComSettingsStr = Ports.ComSettingsStr(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    End Select
10
End Property
