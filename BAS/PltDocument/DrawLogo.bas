Private Sub DrawLogo()
    If lViewLogo And Curves.count = 0 Then
        Dim k As Single
        Dim curvesMatrix As New waMatrix
        k = 700000 / (MaxSheetWidth + 2 * WhiteOffset)
        
        With curvesMatrix
            .ScaleM 1 / mScale / k, -1 / mScale / k
            .Translate OnScreenSheetPositionX + OnScreenSheetWidth / 2, OnScreenSheetPositionY + OnScreenSheetHeight / 2
        End With
        
        myGraphics.SetWorldTransform OnScreenMatrix
        SetOnScreenClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + 1
        Set myGraphics.Curves = GetLogoCurves
        
        Dim P1 As waPath
        Dim P2 As waPath
        Set P1 = myGraphics.Curves(1)
        Set P2 = myGraphics.Curves(2)
        
        Dim c As Long
        c = CurvesPen.ColorARGB
        
        Set P1.pen = CreatePen(, 100, 1.6)
        Set P2.pen = CreatePen(, 100, 2)
        
        P1.pen.ColorARGB = c
        P2.pen.ColorARGB = c
        
        
        myGraphics.SetWorldTransform curvesMatrix
        myGraphics.Render
    End If
End Sub
