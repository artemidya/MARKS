Private Sub GetColor(shape As VGCore.shape)
    Dim c As VGCore.Color
    With shape.Fill
        Dim Color As RGBA
        If .Type = cdrUniformFill Then
            If .UniformColor.Type = cdrColorRGB Then
                Set c = .UniformColor
            Else
                Set c = .UniformColor.GetCopy
                c.ConvertToRGB
            End If
            Color.R = c.RGBRed
            Color.g = c.RGBGreen
            Color.b = c.RGBBlue
            
            'Transparency
            If shape.Transparency.Type = cdrUniformTransparency Then
                Color.a = (100 - shape.Transparency.Uniform) * 2.55
                If Color.a < 127 Then Color.a = 127
            Else
                Color.a = 255
            End If
            
'        Else
'            Color.R = 255
'            Color.G = 255
'            Color.B = 255
'            Color.A = 255
        End If
        PlusColor Color
    End With
End Sub
