Private Function OutType(OutPutName As String) As OutType
    Dim q As Long
    For q = 0 To printers.count - 1
        If printers.Names(q) = OutPutName Then
            OutType = Printer
            Exit Function
        End If
    Next
    
    For q = 0 To Interfaces.count - 1
        If Interfaces.Names(q) = OutPutName Then
            OutType = Interface
            Exit Function
        End If
    Next
    
    For q = 0 To Ports.count - 1
        If Ports.Names(q) = OutPutName Then
            OutType = Port
            Exit Function
        End If
    Next
    
    If OutPutName = "Wi-Fi / Ethernet" Then
        OutType = WiFi
        Exit Function
    End If
        
    '-----
    'FILE
    OutType = file
End Function
