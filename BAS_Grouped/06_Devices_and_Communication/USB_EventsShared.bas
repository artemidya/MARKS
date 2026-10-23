' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: USB_EventsShared.bas
' Procedures: 1
' =============================================================================

' -----------------------------------------------------------------------------
' WindowProc (Function), source lines 20-26
' -----------------------------------------------------------------------------
Public Function WindowProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If uMsg = WM_DEVICECHANGE Then
        lastWorkingTime = Timer
        Changed = True
    End If
    WindowProc = CallWindowProc(oldWndProc, hwnd, uMsg, wParam, lParam)
End Function
