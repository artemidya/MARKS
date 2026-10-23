Private Sub Class_Terminate()
    If Initialized Then
        GdipDeleteGraphics lHGraphics
        GdiplusShutdown lGdipToken
'        ReleaseDC lHWndForm, lhDC
        Initialized = False
    End If
End Sub
