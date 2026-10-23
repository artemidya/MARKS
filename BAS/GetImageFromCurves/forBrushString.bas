Private Function forBrushString(shape As VGCore.shape) As String
    Dim c As Long:    c = 1000
    
    Dim R As Long
    Dim g As Long
    Dim b As Long
    
    Dim R2 As Long
    Dim G2 As Long
    Dim b2 As Long
    
    Dim x1 As Long
    Dim y1 As Long
    Dim x2 As Long
    Dim y2 As Long
    
    forBrushString = vbTab & vbTab & "Set .Brush = new waBrush" & vbNewLine
    
    With shape
        If .Fill.Type = cdrUniformFill Then
            .Fill.UniformColor.ConvertToRGB
            R = .Fill.UniformColor.RGBRed
            g = .Fill.UniformColor.RGBGreen
            b = .Fill.UniformColor.RGBBlue
            forBrushString = forBrushString & vbTab & vbTab & ".Brush.SetColor RGB(" & R & ", " & g & ", " & b & "), 100"

        ElseIf .Fill.Type = cdrFountainFill Then
            .Fill.Fountain.colors.First.Color.ConvertToRGB
            .Fill.Fountain.colors.Last.Color.ConvertToRGB
            
            R = .Fill.Fountain.colors.First.Color.RGBRed
            g = .Fill.Fountain.colors.First.Color.RGBGreen
            b = .Fill.Fountain.colors.First.Color.RGBBlue
            
            R2 = .Fill.Fountain.colors.Last.Color.RGBRed
            G2 = .Fill.Fountain.colors.Last.Color.RGBGreen
            b2 = .Fill.Fountain.colors.Last.Color.RGBBlue
            
            x1 = .Fill.Fountain.startX
            y1 = .Fill.Fountain.startY
            x2 = .Fill.Fountain.EndX
            y2 = .Fill.Fountain.EndY
            
            '.Brush.SetLinearGradient x1, y1, x2, y2, RGB(r, g, b), RGB(r2, g2, b2), WrapModeTileFlipXY
            forBrushString = forBrushString & vbTab & vbTab & ".Brush.SetLinearGradient " & x1 * c & ", " & y1 * c & ", " & x2 * c & ", " & y2 * c & ", RGB(" & R & ", " & g & ", " & b & "), RGB(" & R2 & ", " & G2 & ", " & b2 & "), WrapModeTileFlipXY"
        End If
    End With
End Function
