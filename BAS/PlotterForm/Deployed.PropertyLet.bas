Private Property Let Deployed(oValue As Boolean)
    lDeployed = oValue
    If lDeployed Then
        PlotterForm.Height = 657.75
    Else
        PlotterForm.Height = 525
    End If
End Property
