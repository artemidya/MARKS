Public Sub USB_change_check_loop()
    If USB_EventsShared.Changed And (Timer - USB_EventsShared.lastWorkingTime) > 0.6 Then
        USB_EventsShared.Changed = False
        RaiseEvent Change
    End If
End Sub
