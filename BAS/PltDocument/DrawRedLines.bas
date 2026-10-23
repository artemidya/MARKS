Private Sub DrawRedLines()
    Dim LV As Boolean
    Dim RV As Boolean
    Dim UH As Boolean
    Dim BH As Boolean
    
    Dim dy As Long
    Dim dY2 As Long
    Dim margin As Long
    
    Dim sheetHeight As Long
    If AllCurvesHeight > MaxSheetHeight Then
        sheetHeight = MaxSheetHeight
    Else
        sheetHeight = AllCurvesHeight
    End If
    
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 1, ScreenHeight - 1
    
    Select Case lZeroPosition
        Case RightP
            If AllCurvesWidth + lYmargin > MaxSheetWidth Then LV = True
            If AllCurvesHeight + lXmargin > MaxSheetHeight Then UH = True
            dy = (lXmargin - (MaxSheetHeight - sheetHeight)) / mScale
            margin = lXmargin / mScale
        Case LeftP
            If AllCurvesWidth + lXmargin > MaxSheetWidth Then RV = True
            If AllCurvesHeight + lYmargin > MaxSheetHeight Then UH = True
            dy = (lYmargin - (MaxSheetHeight - sheetHeight)) / mScale
            margin = lYmargin / mScale
        Case CenterP
            If AllCurvesWidth / 2 + lXmargin > MaxSheetWidth / 2 Then RV = True
            If -AllCurvesWidth / 2 + lXmargin < -MaxSheetWidth / 2 Then LV = True
            
            If AllCurvesHeight / 2 + lYmargin > MaxSheetHeight / 2 Then UH = True
            If -AllCurvesHeight / 2 + lYmargin < -MaxSheetHeight / 2 Then BH = True
            
            dy = (lYmargin - (MaxSheetHeight - sheetHeight) / 2) / mScale
            dY2 = (lYmargin + (MaxSheetHeight - sheetHeight) / 2) / mScale
            margin = lYmargin / mScale
            If margin < 0 Then margin = 0
    End Select

    Dim redPen As waPen, whitePen As waPen
    Set redPen = CreatePen(RGB(255, 0, 0), 90, 2)
    Set whitePen = CreatePen(RGB(255, 255, 255), 50, 6)
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    If LV Then
        x1 = OnScreenSheetPositionX + WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    
    If RV Then
        x1 = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    
    If UH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + WhiteOffset / mScale + dy
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        If y1 > OnScreenSheetPositionY Then
            myGraphics.DrawLine whitePen, x1, y1, x2, y2
        End If
    End If
    
    If BH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale + dY2
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    'Red===================================================
    If LV Then
        x1 = OnScreenSheetPositionX + WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    
    If RV Then
        x1 = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    
    If UH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + WhiteOffset / mScale + dy
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        If y1 > OnScreenSheetPositionY Then
            myGraphics.DrawLine redPen, x1, y1, x2, y2
        End If
    End If
    
    If BH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale + dY2
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    lRedLineError = LV Or RV Or UH Or BH
End Sub
