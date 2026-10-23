Public Sub UserFormRefreshData()
    On Error Resume Next
    Dim k As Double, errStr As String
    Dim Area As Double
    
    If CountErr > 0 And ComboBox1.ListIndex <> 2 Then
        errStr = "err(" & CalculateSanP.CountErr & ") "
        TextBox3.ForeColor = RGB(250, 0, 0)
        TextBox3.ControlTipText = localization.ActiveLng.Item("SquareForm.TextBox3.ControlTipText") & " (" & CalculateSanP.CountErr & " " & localization.ActiveLng.Item("SquareForm.pcs") & ")"
    Else
        errStr = ""
        TextBox3.ForeColor = TextBox2.ForeColor
        TextBox3.ControlTipText = ""
    End If

    Dim k2 As Double: k2 = VGCore.Application.ConvertUnits(1, CalculateSanP.lastUnit, cdrInch)
    k = 0.00064516 * k2 * k2
    
    If ComboBox1.ListIndex = 1 Then
        TextBox3 = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.RectAreaOfCurves * k, "0.000000")
        TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.RectAreaOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        CalculateSanP.ErrShow
    ElseIf ComboBox1.ListIndex = 2 Then
        TextBox3 = Format(CalculateSanP.SelectionRectArea * k, "0.000000")
        TextBox5.text = Format(CalculateSanP.SelectionRectArea * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        If CalculateSanP.PerimetrFormVisible Then
            CalculateSanP.ErrShow
        Else
            CalculateSanP.ErrHide
        End If
    Else
        TextBox3 = errStr & CalculateSanP.ProcessingStr & Format(CalculateSanP.AreaOfCurves * k, "0.000000")
        TextBox5.text = CalculateSanP.ProcessingStr & Format(CalculateSanP.AreaOfCurves * k * ValP(TextBox1.text) * (100 + ValP(TextBox2)) / 100, "0.00")
        CalculateSanP.ErrShow
    End If
    
End Sub
