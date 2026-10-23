Private Sub DrawSheetShadow()
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
        myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
        myGraphics.FillRect CreateBrush(RGB(170, 170, 170), 100), OnScreenSheetPositionX + lShadowDistance, OnScreenSheetPositionY + lShadowDistance, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
    End If
End Sub
