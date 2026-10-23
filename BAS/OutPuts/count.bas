Public Property Get count() As Long
    count = printers.count + Interfaces.count + Ports.count + 2
End Property
