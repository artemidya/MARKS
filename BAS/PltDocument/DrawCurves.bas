Private Sub DrawCurves()
    Dim curvesMatrix As New waMatrix
    
    With curvesMatrix
        .ScaleM 1 / mScale, -1 / mScale
        Select Case lZeroPosition
            Case RightP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale - lYmargin / mScale, OnScreenOrignY
            Case LeftP
                .Translate OnScreenOrignX + lXmargin / mScale, OnScreenOrignY
            Case CenterP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale / 2 + lXmargin / mScale, OnScreenOrignY + AllCurvesHeight / mScale / 2
        End Select
    End With
    
    Dim hPlus As Long
    If lZeroPosition = RightP Then
        hPlus = lXmargin / mScale
    Else
        If lZeroPosition = CenterP Then
            If lYmargin < 0 Then
                hPlus = 0
            Else
                hPlus = lYmargin / mScale
            End If
        Else
            hPlus = lYmargin / mScale
        End If
    End If
    
    SetOnScreenClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + 1 + hPlus
    Set myGraphics.Curves = Curves
    myGraphics.SetWorldTransform curvesMatrix
    myGraphics.Render , , , CurvesPen
End Sub
