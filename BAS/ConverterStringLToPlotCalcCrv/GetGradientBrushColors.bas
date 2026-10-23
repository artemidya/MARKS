Private Function GetGradientBrushColors(ByRef StringL As String) As GradientBrushParametrs
    GetGradientBrushColors.Color1 = GetSolidBrushColor(StringL)
    GetGradientBrushColors.Color2 = GetSolidBrushColor(StringL)
    GetGradientBrushColors.x1 = GetCoordinate(StringL)
    GetGradientBrushColors.y1 = GetCoordinate(StringL)
    GetGradientBrushColors.x2 = GetCoordinate(StringL)
    GetGradientBrushColors.y2 = GetCoordinate(StringL)
End Function
