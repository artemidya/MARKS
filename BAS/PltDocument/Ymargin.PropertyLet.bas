Property Let Ymargin(oValue As Long)
    If lYmargin <> oValue Then
        lYmargin = oValue
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property
