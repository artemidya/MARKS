Private Function MaxColorF() As RGBA

    Dim colorInfo1 As ColorInfo
    Dim MaxColorInfo As New ColorInfo
    
    MaxColorInfo.R = 255
    MaxColorInfo.g = 255
    MaxColorInfo.b = 255
    MaxColorInfo.a = 255
    
    For Each colorInfo1 In MyColors
        If colorInfo1.count > MaxColorInfo.count And colorInfo1.a <> 0 Then
            Set MaxColorInfo = colorInfo1
        End If
    Next
    
    MaxColorF.R = MaxColorInfo.R
    MaxColorF.g = MaxColorInfo.g
    MaxColorF.b = MaxColorInfo.b
    MaxColorF.a = MaxColorInfo.a
End Function
