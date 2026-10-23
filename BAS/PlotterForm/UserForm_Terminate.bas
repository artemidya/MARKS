Private Sub UserForm_Terminate()
    KeyLoopProcessing = False
    If Not USB_Events1 Is Nothing Then USB_Events1.Terminate
    If Not BlocksSocket1 Is Nothing Then BlocksSocket1.Disconnect
    Set OutPuts = Nothing
    Dim q As Long
    For q = 0 To CountProfiles - 1
        Set Profiles(q) = Nothing
    Next
    Converter.Break
    StopShowAll
End Sub
