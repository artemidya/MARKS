Private Function ListPrintersStatus() As ConnectionStatus()
    On Error GoTo 10
    
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object

    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer")

    If colItems.count > 0 Then
        Dim ConnectionStatuses() As ConnectionStatus
        ReDim ConnectionStatuses(colItems.count - 1)
    
        Dim i As Long
        
        For Each objItem In colItems
            With ConnectionStatuses(i)
    '            Debug.Print objItem.Name & " | Port: " & objItem.portName & " | Status: " & objItem.WorkOffline
                .name = objItem.name
                .Connected = Not objItem.WorkOffline
                i = i + 1
            End With
        Next
        ListPrintersStatus = ConnectionStatuses
    End If
10
End Function
