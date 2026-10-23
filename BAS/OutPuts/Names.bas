Public Property Get Names(ByVal i As Long) As String
    If i < printers.count Then
        Names = printers.Names(i)
    ElseIf i < (printers.count + Interfaces.count) Then
        i = i - printers.count
        Names = Interfaces.Names(i)
    ElseIf i < (printers.count + Interfaces.count + Ports.count) Then
        i = i - printers.count - Interfaces.count
        Names = Ports.Names(i)
    ElseIf i = (printers.count + Interfaces.count + Ports.count) Then
        Names = "Wi-Fi / Ethernet"
    Else
        Names = "File (flash card)"
    End If
End Property
