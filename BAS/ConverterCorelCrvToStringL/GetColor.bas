Private Function GetColor(shape As VGCore.shape) As String
    Dim c As VGCore.Color
    With shape.Fill
        If .Type = cdrUniformFill Then
            'Solid
            If .UniformColor.Type = cdrColorRGB Then
                Set c = .UniformColor
            Else
                Set c = .UniformColor.GetCopy
                c.ConvertToRGB
            End If
            
            GetColor = "S" & c.RGBRed & "," & c.RGBGreen & "," & c.RGBBlue
              
            'Transparency
            If shape.Transparency.Type = cdrUniformTransparency Then
                 GetColor = GetColor & "," & Int((100 - shape.Transparency.Uniform) * 2.55)
            Else
                 GetColor = GetColor & ",255"
            End If
        ElseIf .Type = cdrFountainFill And .Fountain.Type = cdrLinearFountainFill Then
            'Linear Gradient
            If .Fountain.colors.First.Color.Type = cdrColorRGB Then
                Set c = .Fountain.colors.First.Color
            Else
                Set c = .Fountain.colors.First.Color.GetCopy
                c.ConvertToRGB
            End If
            
            Dim c2 As VGCore.Color
            If .Fountain.colors.Last.Color.Type = cdrColorRGB Then
                Set c2 = .Fountain.colors.Last.Color
            Else
                Set c2 = .Fountain.colors.Last.Color.GetCopy
                c2.ConvertToRGB
            End If
            
            'Transparency
            Dim TR As String
            If shape.Transparency.Type = cdrUniformTransparency Then
                 TR = "," & Int((100 - shape.Transparency.Uniform) * 2.55)
            Else
                 TR = ",255"
            End If
            
            GetColor = "G" & c.RGBRed & "," & c.RGBGreen & "," & c.RGBBlue & TR & "," _
                           & c2.RGBRed & "," & c2.RGBGreen & "," & c2.RGBBlue & TR & "," _
                           & CLng(.Fountain.startX * 1000) & "," & CLng(.Fountain.startY * 1000) & "," _
                           & CLng(.Fountain.EndX * 1000) & "," & CLng(.Fountain.EndY * 1000)
        Else
            GetColor = "S255,255,255,255"
        End If
    End With
End Function
