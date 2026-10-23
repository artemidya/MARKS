' =============================================================================
' Group: 02_Calculation_and_Nesting_Forms
' Source module: ProgressBar.cls
' Procedures: 3
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 14-18
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    RaskroyForm.PBar.Caption = ""
    RaskroyForm.PBarStrip.Width = 0
    k = RaskroyForm.PBar.Width / 10000
End Sub

' -----------------------------------------------------------------------------
' value (PropertyLet), source lines 19-41
' -----------------------------------------------------------------------------
Property Let value(oValue As Long)
    RaskroyForm.PBarStrip.Width = k * oValue
    If oValue > 0 Then
        If NeedDoubleDigits And oValue <> 10000 Then
            RaskroyForm.PBar.Caption = Format(oValue / 100, "0.00") & "%"
        Else
            RaskroyForm.PBar.Caption = oValue \ 100 & "%"
        End If
    Else
        RaskroyForm.PBar.Caption = ""
    End If
    
    If Not NeedDoubleDigits Then
        If oValue <> 0 Then
            If t <> 0 Then
                If Timer - t > 1 Then NeedDoubleDigits = True
            End If
            t = Timer
        Else
            t = 0
        End If
    End If
End Property

' -----------------------------------------------------------------------------
' GreenColorOn (PropertyLet), source lines 42-48
' -----------------------------------------------------------------------------
Property Let GreenColorOn(oGreenColorOn As Boolean)
    If oGreenColorOn Then
        RaskroyForm.PBarStrip.BackColor = RGB(147, 244, 186)
    Else
        RaskroyForm.PBarStrip.BackColor = VBA.SystemColorConstants.vbActiveTitleBar
    End If
End Property
