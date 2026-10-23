Private Function GetHostByNameAlias(ByVal hostname As String) As Long
    On Error Resume Next
    Dim phe As LongPtr
    Dim heDestHost As HostEnt
    Dim addrList As LongPtr
    Dim retIP As Long
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            MemMove heDestHost, ByVal phe, LenB(heDestHost)
            MemMove addrList, ByVal heDestHost.h_addr_list, LenB(heDestHost.h_addr_list)
            MemMove retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function
