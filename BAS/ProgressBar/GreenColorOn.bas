Property Let GreenColorOn(oGreenColorOn As Boolean)
    If oGreenColorOn Then
        RaskroyForm.PBarStrip.BackColor = RGB(147, 244, 186)
    Else
        RaskroyForm.PBarStrip.BackColor = VBA.SystemColorConstants.vbActiveTitleBar
    End If
End Property
