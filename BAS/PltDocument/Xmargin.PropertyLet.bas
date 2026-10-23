Property Let Xmargin(oValue As Long)
    If lXmargin <> oValue Then
        lXmargin = oValue
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property
