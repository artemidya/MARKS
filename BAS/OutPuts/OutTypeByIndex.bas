Property Get OutTypeByIndex(ByVal index As Long) As OutType
    If index < printers.count Then
        OutTypeByIndex = Printer
    ElseIf index < (printers.count + Interfaces.count) Then
        OutTypeByIndex = Interface
    ElseIf index < (printers.count + Interfaces.count + Ports.count) Then
        OutTypeByIndex = Port
    ElseIf index < (printers.count + Interfaces.count + Ports.count + 1) Then
        OutTypeByIndex = WiFi
    ElseIf index < (printers.count + Interfaces.count + Ports.count + 2) Then
        OutTypeByIndex = file
    End If
End Property
