' =============================================================================
' Group: 01_Application_and_Menu
' Source module: MonitorDPI.bas
' Procedures: 2
' =============================================================================

' -----------------------------------------------------------------------------
' GetDpiDbl (Function), source lines 18-51
' -----------------------------------------------------------------------------
Public Function GetDpiDbl() As Double
    Dim dpi As Long
    dpi = GetDpi
    '96 – Smaller 100%
    '120 – Medium 125%
    '144 – Larger 150%
    '192 – Extra Large 200%
    '240 – Custom 250%
    '288 – Custom 300%
    '384 – Custom 400%
    '480 – Custom 500%
    GetDpiDbl = dpi / 96

'    Select Case dpi
'        Case 98
'            GetDpiDbl = 1
'        Case 120
'            GetDpiDbl = 1.25
'        Case 144
'            GetDpiDbl = 1.5
'        Case 192
'            GetDpiDbl = 2
'        Case 240
'            GetDpiDbl = 2.5
'        Case 288
'            GetDpiDbl = 3
'        Case 384
'            GetDpiDbl = 4
'        Case 480
'            GetDpiDbl = 5
'        Case Else
'            GetDpiDbl = 1
'    End Select
End Function

' -----------------------------------------------------------------------------
' GetDpi (Function), source lines 52-76
' -----------------------------------------------------------------------------
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
