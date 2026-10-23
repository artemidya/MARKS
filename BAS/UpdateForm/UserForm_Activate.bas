Private Sub UserForm_Activate()
    On Error GoTo 10
    Set UpdateCenter1 = New UpdateCenter
    If UpdateCenter1.version > UserData.PlotCalcVersion Then
        UpdateCenter1.UpdatePlotCalc
    End If
    Unload Me
10
End Sub
