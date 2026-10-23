Public Function GetDpi() As Long
    Dim iDPI As Long
    iDPI = 98
    On Error GoTo 10
    
    #If VBA7 Then
        Dim hdcScreen As LongPtr
        Dim hwnd As LongPtr
    #Else
        Dim hdcScreen As Long
        Dim hwnd As Long
    #End If
    
    hwnd = GetActiveWindow()
    hdcScreen = GetDC(hwnd)



    If (hdcScreen) Then
        iDPI = GetDeviceCaps(hdcScreen, LOGPIXELSX)
        ReleaseDC hwnd, hdcScreen
    End If
10
    GetDpi = iDPI
End Function
