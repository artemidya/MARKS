Private Function GetBrush(ByRef StringL As String) As waBrush
    For cp = cp To LenStringL
        If IsBrush(StringL) Then
            If Mid(StringL, cp, 1) = "S" Then
                'Solid
                Dim col As RGBA
                col = GetSolidBrushColor(StringL)
                Set GetBrush = CreateBrush(RGB(col.R, col.g, col.b), CLng(CDbl(col.a) / 255 * 100))
            Else
                'Gradient
                Dim gbp As GradientBrushParametrs
                gbp = GetGradientBrushColors(StringL)
                Set GetBrush = New waBrush
                With gbp
                    GetBrush.SetLinearGradient .x1, .y1, .x2, .y2, RGB(.Color1.R, .Color1.g, .Color1.b), RGB(.Color2.R, .Color2.g, .Color2.b), , CLng(CDbl(.Color1.a) / 255 * 100), CLng(CDbl(.Color2.a) / 255 * 100)
                End With
            End If
            Exit Function
        End If
    Next
End Function
