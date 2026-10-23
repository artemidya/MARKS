Private Sub DrawSheet()
    If Curves.count > 0 Or lViewLogo Then
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
    
        myGraphics.SetWorldTransform OnScreenMatrix
        myGraphics.SetClip 0, 0, ScreenWidth - 1, ScreenHeight - 1
        
        myGraphics.FillRect lMaterialBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
        
'        Dim gradientBrush As New waBrush
'        gradientBrush.SetLinearGradient OnScreenSheetPositionX, OnScreenSheetPositionY + WhiteOffset / mScale + 1, OnScreenSheetPositionX, OnScreenSheetPositionY, 0, 0, , 0, 30
        
'        myGraphics.SetClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + hPlus + 1
'        myGraphics.FillRect gradientBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, WhiteOffset / mScale
'        myGraphics.SetClip 0, 0, ScreenWidth - 1, ScreenHeight - 1
        myGraphics.DrawRect CreatePen(RGB(150, 150, 150), 100), OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
    End If
End Sub
