Private Sub UnHookForm()  '<<<< ------------------ Private
    If IsHooked = True Then
        If PrevProc <> 0 Then
            SetWindowLong Hooked_hWnd, GWL_WNDPROC, PrevProc
            PrevProc = 0
        End If
        IsHooked = False
    End If
End Sub
