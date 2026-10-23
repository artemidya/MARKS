Public Property Get Xmm() As Double
    Dim mOutType As OutType
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Dim x As Double
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
    
    If Profiles(PlotterForm.TabStrip1.SelectedItem.index).ZeroPosition = RightP Then
        x = x / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).HeightScale) / 40
    Else
        x = x / Abs(Profiles(PlotterForm.TabStrip1.SelectedItem.index).WidthScale) / 40
    End If
    
    Xmm = x
End Property
