Public Property Get x() As Long
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            x = printers.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            x = Interfaces.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            x = Ports.x(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            x = WiFiCoord.x
        Case file
            x = FileCoord.x
    End Select
End Property
