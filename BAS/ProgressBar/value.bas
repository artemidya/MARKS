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
