Property Let XYmargin(oValue As POINTL)
    If lXmargin <> oValue.x Or lYmargin <> oValue.y Then
        lXmargin = oValue.x
        lYmargin = oValue.y
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property
