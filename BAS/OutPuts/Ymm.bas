Public Property Get Ymm() As Double
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Dim y As Double
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
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        y = y / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale) / 40
    Else
        y = y / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale) / 40
    End If
    
    Ymm = y
End Property
