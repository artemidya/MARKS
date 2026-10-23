Private Sub DrawAxis()
    If lViewAxis Then
        Dim hPlus As Long
        If lZeroPosition = RightP Then
            hPlus = lXmargin / mScale
        Else
            hPlus = lYmargin / mScale
        End If
            
        myGraphics.SetWorldTransform OnScreenMatrix
        myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
        
        Set myGraphics.Curves = AxisCurves
        Dim myMatrix As New waMatrix
        myMatrix.ScaleM 0.002, -0.002
        
        If lZeroPosition = RightP Then myMatrix.Rotate -90
        myMatrix.Translate OnScreenOrignX, OnScreenOrignY + hPlus
        myGraphics.SetWorldTransform myMatrix
        
        myGraphics.Render False, True
    End If
End Sub
