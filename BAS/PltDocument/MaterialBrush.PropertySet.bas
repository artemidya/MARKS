Property Set MaterialBrush(oValue As waBrush)
    Set lMaterialBrush = oValue
    
    Dim b As Long
    b = lMaterialBrush.Brightness
    
    If b < 150 Then
        'Dark Material
        Set CurvesPen = CreatePen(RGB(210, 210, 210), 80, 1.6)
        Set WayPen = CreatePen(RGB(255, 237, 0), 50, 2)
        Set UpWayPen = CreatePen(RGB(255, 237, 0), 25, 2)
        
    Else
        'Light Material
        Set CurvesPen = CreatePen(RGB(110, 110, 110), 80, 1.6)
        Set WayPen = CreatePen(RGB(57, 49, 133), 50, 2)
        Set UpWayPen = CreatePen(RGB(57, 49, 133), 30, 2)
    End If
    
    UpWayPen.DashStyle = Dash
    UpWayPen.EndCap = ArrowAnchor
    WayPen.EndCap = ArrowAnchor
    
    Dim m As New waMatrix
    m.ScaleM 0.025, 0.025
    Set WayPen.matrix = m
    Set UpWayPen.matrix = m
        
End Property
