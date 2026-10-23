Private Sub Class_Terminate()
    If notifyHandle <> 0 Then
        UnregisterDeviceNotification notifyHandle
    End If
    If notifyHandle2 <> 0 Then
        UnregisterDeviceNotification notifyHandle2
    End If
    UnhookWindow
End Sub
