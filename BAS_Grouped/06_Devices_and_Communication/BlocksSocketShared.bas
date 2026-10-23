' =============================================================================
' Group: 06_Devices_and_Communication
' Source module: BlocksSocketShared.bas
' Procedures: 5
' =============================================================================

' -----------------------------------------------------------------------------
' HookForm (Sub), source lines 52-61
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' UnHookForm (Sub), source lines 62-70
' -----------------------------------------------------------------------------
Private Sub UnHookForm()  '<<<< ------------------ Private
    If IsHooked = True Then
        If PrevProc <> 0 Then
            SetWindowLong Hooked_hWnd, GWL_WNDPROC, PrevProc
            PrevProc = 0
        End If
        IsHooked = False
    End If
End Sub

' -----------------------------------------------------------------------------
' StartWinsock (Function), source lines 72-83
' -----------------------------------------------------------------------------
Public Function StartWinsock(Optional sDescription) As Boolean
    If IsHooked = False Then MsgBox "No Hooked Form! Error": Exit Function
    CountCreationSockets = CountCreationSockets + 1
    If Not WSAStartedUp Then
        Dim StartupData As WSADataType
        If Not WSAStartup(&H101, StartupData) Then
            WSAStartedUp = True
            sDescription = StartupData.szDescription
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

' -----------------------------------------------------------------------------
' EndWinsock (Sub), source lines 85-97
' -----------------------------------------------------------------------------
Public Sub EndWinsock()
    CountCreationSockets = CountCreationSockets - 1
    If CountCreationSockets > 0 Then Exit Sub
    If CountCreationSockets < 0 Then CountCreationSockets = 0

    Dim ret As Long
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
    UnHookForm
End Sub

' -----------------------------------------------------------------------------
' WindowProc (Function), source lines 98-107
' -----------------------------------------------------------------------------
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
