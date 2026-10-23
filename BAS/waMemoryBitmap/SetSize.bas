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
