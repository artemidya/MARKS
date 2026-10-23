Public Property Get Connected(ByVal i As Long) As Boolean
    If i < printers.count Then
        Connected = printers.Connected(i)
    ElseIf i < (printers.count + Interfaces.count) Then
        i = i - printers.count
        Connected = Interfaces.Connected(i)
    ElseIf i < (printers.count + Interfaces.count + Ports.count) Then
        i = i - printers.count - Interfaces.count
        Connected = Ports.Connected(i)
    ElseIf i = (printers.count + Interfaces.count + Ports.count) Then
        
        Connected = PlotterClientConnected '
    Else
        Connected = FileConnectedIfUSB ' "File (flash card)"
    End If

End Property
