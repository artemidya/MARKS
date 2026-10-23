' =============================================================================
' Group: 05_Graphics_Engine
' Source module: waMemoryBitmap.cls
' Procedures: 5
' =============================================================================

' -----------------------------------------------------------------------------
' hDC (PropertyGet), source lines 26-28
' -----------------------------------------------------------------------------
Property Get hDC() As LongPtr
    hDC = lhDC
End Property

' -----------------------------------------------------------------------------
' Height (PropertyGet), source lines 29-31
' -----------------------------------------------------------------------------
Property Get Height() As Long
    Height = lHeight
End Property

' -----------------------------------------------------------------------------
' Width (PropertyGet), source lines 32-34
' -----------------------------------------------------------------------------
Property Get Width() As Long
    Width = lWidth
End Property

' -----------------------------------------------------------------------------
' SetSize (Sub), source lines 35-58
' -----------------------------------------------------------------------------
Public Sub SetSize(Width As Long, Height As Long)
    If Not SizeIsSet Then
        'Init
        Dim hwnd As LongPtr
        Dim hDCfrom As LongPtr
        
        hwnd = GetActiveWindow
        hDCfrom = GetDC(hwnd)
        
        ' Create the device context.
        lhDC = CreateCompatibleDC(hDCfrom)
        ' Create the bitmap.
        hBM = CreateCompatibleBitmap(hDCfrom, Width, Height)
        ReleaseDC hwnd, hDCfrom
        
        ' Make the device context use the bitmap.
        oldhDC = SelectObject(lhDC, hBM)
        lWidth = Width
        lHeight = Height
        SizeIsSet = True
    Else
        MsgBox "Size already set", vbCritical
    End If
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 60-66
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
     If SizeIsSet Then
        SelectObject lhDC, oldhDC
        DeleteObject hBM
        DeleteDC lhDC
     End If
End Sub
