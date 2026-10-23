Public Function Out(ByRef bytes() As Byte, x As Long, y As Long, Optional dntInitialize As Boolean = False) As Boolean
    On Error GoTo 10
    Dim mOutType As OutType
    Dim f As Boolean
    mOutType = OutType(KillBirdsAndSpaces(PlotterForm.Combo1.text))
    Select Case mOutType
        Case Printer
            printers.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y
            Out = True
        Case Interface
            Interfaces.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y, dntInitialize
            Out = True
        Case Port
            Ports.Out KillBirdsAndSpaces(PlotterForm.Combo1.text), bytes, x, y, dntInitialize
            Out = True
        Case WiFi
            WiFiCoord.x = x
            WiFiCoord.y = y
            OutToWiFi bytes()
            Out = True
        Case file
            FileCoord.x = x
            FileCoord.y = y
            f = OutToFile(bytes)
            Out = f
    End Select
    
    If (mOutType = file And f) Or mOutType <> file Then
        If UBound(bytes) < 7000 Then
            PlotterForm.AddLog PlotterLanguage.AsciiToString(bytes)
        Else
            PlotterForm.AddLog "Write " & UBound(bytes) + 1 & " bytes"
        End If
    End If
10
End Function
