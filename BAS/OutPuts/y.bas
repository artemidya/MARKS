Public Property Get y() As Long
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            y = printers.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Interface
            y = Interfaces.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case Port
            y = Ports.y(KillBirdsAndSpaces(PlotterForm.Combo1.text))
        Case WiFi
            y = WiFiCoord.y
        Case file
            y = FileCoord.y
    End Select
End Property
