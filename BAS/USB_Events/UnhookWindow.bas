Private Sub UnhookWindow()
    If oldWndProc <> 0 Then
        SetWindowLongPtr hWndForm, GWL_WNDPROC, oldWndProc
    End If
End Sub
