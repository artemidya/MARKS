Private Function IsCOMPortAvailable(ByVal PortName As String) As Boolean
    On Error GoTo 10
    PortName = Trim(PortName)
    PortName = RemoveTrailingColon(PortName)
    Dim Handle As LongPtr
    Handle = CreateFileA("\\.\" & PortName, GENERIC_READ, 0, 0, OPEN_EXISTING, 0, 0)
    
    If Handle = INVALID_HANDLE_VALUE Then
        IsCOMPortAvailable = False
    Else
        IsCOMPortAvailable = True
        CloseHandle Handle
    End If
10
End Function
