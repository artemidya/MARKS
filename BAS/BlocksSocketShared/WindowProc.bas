Private Function WindowProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    'MsgBox hWnd
    On Error Resume Next
    If uMsg = WINSOCKMSG Then
        'our Winsock-message handler
        OpenedSockets(CStr(wParam)).ProcessMessage lParam
    Else
        WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    End If
End Function
