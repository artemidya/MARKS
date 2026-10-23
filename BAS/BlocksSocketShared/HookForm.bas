Public Sub HookForm(ByVal hwnd As LongPtr)
    If IsHooked = False Then
        Hooked_hWnd = hwnd
        PrevProc = SetWindowLong(Hooked_hWnd, GWL_WNDPROC, AddressOf WindowProc)
'        If PrevProc = 0 Then
'            MsgBox err.LastDllError
'        End If
        IsHooked = True
    End If
End Sub
