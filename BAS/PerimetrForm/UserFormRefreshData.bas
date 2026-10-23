Public Sub UserFormRefreshData()
    On Error Resume Next
    Dim k As Double, errStr As String
    If CountErr > 0 Then
        errStr = "err(" & CalculateSanP.CountErr & ") "
        TextBox3.ForeColor = RGB(250, 0, 0)
        TextBox3.ControlTipText = localization.ActiveLng.Item("PerimetrForm.ImpossibleMeasureLengthOfCurve") & " (" & CalculateSanP.CountErr & " " & localization.ActiveLng.Item("PerimetrForm.unit") & ")"
        CalculateSanP.ErrShow
    Else
        errStr = ""
        TextBox3.ForeColor = TextBox2.ForeColor
        TextBox3.ControlTipText = ""
    End If
    
    k = 0.0254 * VGCore.Application.ConvertUnits(1, CalculateSanP.lastUnit, cdrInch)
    TextBox3.text = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.LengthOfCurves * k, "0.000")
    TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.LengthOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
    TextBox6.text = CalculateSanP.ProcessingStr & CalculateSanP.CountSubPaths
    TextBox7.text = CalculateSanP.ProcessingStr & CalculateSanP.CountNodes
End Sub
