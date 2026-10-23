Private Sub HookWindow()
    oldWndProc = GetWindowLongPtr(hWndForm, GWL_WNDPROC)
    SetWindowLongPtr hWndForm, GWL_WNDPROC, AddressOf USB_EventsShared.WindowProc
End Sub
