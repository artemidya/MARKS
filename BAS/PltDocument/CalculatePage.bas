Private Sub CalculatePage()
    Dim rScreenW As Long, rScreenH As Long
    rScreenW = ScreenWidth - FromFrameOffset - FromFrameOffset
    rScreenH = ScreenHeight - FromFrameOffset - FromFrameOffset
    
    If Curves.count = 0 Then
        If ActivePlotter.Height / rScreenH > ActivePlotter.Width / rScreenW Then
            'Vertical
            mScale = CDbl(ActivePlotter.Height) / rScreenH
            If lViewLogo Then
                OnScreenSheetHeight = (ActivePlotter.Height * 0.3 + 2 * WhiteOffset) / mScale
            Else
                OnScreenSheetHeight = WhiteOffset / mScale
            End If
            OnScreenSheetWidth = (MaxSheetWidth + WhiteOffset + WhiteOffset) / mScale
                       
            OnScreenSheetPositionX = ScreenWidth / 2 + (Abs(ActivePlotter.left) - MaxSheetWidth - WhiteOffset - ActivePlotter.Width / 2) / mScale
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        Else
            'Horisontal
            mScale = CDbl(ActivePlotter.Width) / rScreenW
            If lViewLogo Then
                OnScreenSheetHeight = (ActivePlotter.Height * 0.3 + 2 * WhiteOffset) / mScale
            Else
                OnScreenSheetHeight = WhiteOffset / mScale
            End If
            OnScreenSheetWidth = (MaxSheetWidth + WhiteOffset + WhiteOffset) / mScale
            
            OnScreenSheetPositionX = FromFrameOffset + (Abs(ActivePlotter.left) - MaxSheetWidth - WhiteOffset) / mScale
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        End If
    Else
        Dim sheetHeight As Long
        If AllCurvesHeight > MaxSheetHeight Then
            sheetHeight = MaxSheetHeight
        Else
            sheetHeight = AllCurvesHeight
        End If
    
        If (sheetHeight + ActivePlotter.top + 2 * WhiteOffset) / rScreenH > (MaxSheetWidth + 2 * WhiteOffset) / rScreenW Then
            'Vertical
            mScale = CDbl(sheetHeight + ActivePlotter.top + 2 * WhiteOffset) / rScreenH
            OnScreenSheetHeight = rScreenH - ActivePlotter.top / mScale
            OnScreenSheetWidth = (MaxSheetWidth + 2 * WhiteOffset) / mScale
            
            OnScreenSheetPositionX = ScreenWidth / 2 - OnScreenSheetWidth / 2
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        Else
            'Horisontal
            mScale = CDbl(MaxSheetWidth + 2 * WhiteOffset) / rScreenW
            OnScreenSheetHeight = (sheetHeight + 2 * WhiteOffset) / mScale
            OnScreenSheetWidth = rScreenW
            
            OnScreenSheetPositionX = FromFrameOffset
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        End If
    End If
    
    Select Case lZeroPosition
        Case RightP
            OnScreenOrignX = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale
        Case LeftP
            OnScreenOrignX = OnScreenSheetPositionX + WhiteOffset / mScale
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale
        Case CenterP
            OnScreenOrignX = OnScreenSheetPositionX + OnScreenSheetWidth / 2
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight / 2
    End Select
End Sub
