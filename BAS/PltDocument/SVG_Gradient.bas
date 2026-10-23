Private Function SVG_Gradient(Curves As Collection, ByRef GNumber As Long) As String
    Dim curve As waPath
    Dim SVGtxt As String
    For Each curve In Curves
        With curve.brush
            If .FillType = BrushTypeLinearGradient Then
                SVGtxt = SVGtxt & "<linearGradient id=" & Chr(34) & "G" & GNumber & Chr(34) & " x1=" & Chr(34) & .P1_GradientLine.x & Chr(34) & " x2=" & Chr(34) & .P2_GradientLine.x & Chr(34) & " y1=" & Chr(34) & -.P1_GradientLine.y & Chr(34) & " y2=" & Chr(34) & -.P2_GradientLine.y & Chr(34) & " spreadMethod=" & Chr(34) & "reflect" & Chr(34) & " gradientUnits=" & Chr(34) & "userSpaceOnUse" & Chr(34) & ">" & Chr(13)
                SVGtxt = SVGtxt & "<stop offset=" & Chr(34) & "0%" & Chr(34) & " stop-color=" & Chr(34) & rgbaToStr(.ColorLineGradient1) & Chr(34) & "/>" & Chr(13)
                SVGtxt = SVGtxt & "<stop offset=" & Chr(34) & "100%" & Chr(34) & " stop-color=" & Chr(34) & rgbaToStr(.ColorLineGradient2) & Chr(34) & "/>" & Chr(13)
                SVGtxt = SVGtxt & "</linearGradient>" & Chr(13)
                GNumber = GNumber + 1
            End If
        End With
    Next
    SVG_Gradient = SVGtxt
End Function
