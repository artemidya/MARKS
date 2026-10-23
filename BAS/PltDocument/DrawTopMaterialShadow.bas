Private Sub DrawTopMaterialShadow()
    If Curves.count > 0 Or lViewLogo Then
        Dim gradientBrush As New waBrush
        gradientBrush.SetLinearGradient OnScreenSheetPositionX, OnScreenSheetPositionY + WhiteOffset / mScale + 1, OnScreenSheetPositionX, OnScreenSheetPositionY, 0, 0, , 0, 30
        myGraphics.SetClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight
        myGraphics.FillRect gradientBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, WhiteOffset / mScale
    End If
End Sub
