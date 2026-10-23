Private Sub Class_Initialize()
    BlocksSocketShared.HookForm GetActiveWindow
    Set socket1 = New BlocksSocket
    GettedVersion = -1
End Sub
