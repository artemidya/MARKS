Private Sub Class_Initialize()
    HorisontalDelitel = 2
    DelayMultiplier = 1
    lHWndForm = GetActiveWindow
    hDC = GetDC(lHWndForm)
    Set myGraphics = New waGraphics
    Set AxisCurves = GetAxisCurves
    lViewAxis = True
    Set lBackgroundBrush = CreateBrush(ColorFromOle(SystemColorConstants.vbScrollBars))
    Set MaterialBrush = CreateBrush(RGB(255, 255, 255))
    
    lShadowDistance = 5
    lViewBorder = True
    lBorderColor = ColorFromOle(SystemColorConstants.vbWindowFrame)
End Sub
