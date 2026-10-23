Property Get SVG_Generator() As String
    Dim GNumber As Long
    Dim AnimationTxt As String
    GNumber = 1
    
    SVG_Generator = "<!-- Creator: plotcalc.com 1000$ -->" & Chr(13)

    With ActivePlotter
        SVG_Generator = SVG_Generator & "<svg version=" & Chr(34) & "1.2" & Chr(34) & " xmlns=" & Chr(34) & "http://www.w3.org/2000/svg" & Chr(34) & " viewBox=" & Chr(34) & .left & " " & -.top & " " & .Width & " " & .Height & Chr(34) & " xmlns:xlink=" & Chr(34) & "http://www.w3.org/1999/xlink" & Chr(34) & ">" & Chr(13)
    End With
 
    'Gradients
    SVG_Generator = SVG_Generator & "<defs>" & Chr(13)
    SVG_Generator = SVG_Generator & SVG_Gradient(ActivePlotter.Curves, GNumber)
    SVG_Generator = SVG_Generator & SVG_Gradient(ActivePlotter.CarriageCurves, GNumber)
    
    SVG_Generator = SVG_Generator & "<linearGradient id=" & Chr(34) & "GS1" & Chr(34) & " x1=" & Chr(34) & "0" & Chr(34) & " x2=" & Chr(34) & "0" & Chr(34) & " y1=" & Chr(34) & "0" & Chr(34) & " y2=" & Chr(34) & "25000" & Chr(34) & " gradientUnits=" & Chr(34) & "userSpaceOnUse" & Chr(34) & ">" & Chr(13)
    SVG_Generator = SVG_Generator & "<stop offset=" & Chr(34) & "25%" & Chr(34) & " stop-color=" & Chr(34) & "rgba(0,0,0,.2)" & Chr(34) & "/>" & Chr(13)
    SVG_Generator = SVG_Generator & "<stop offset=" & Chr(34) & "100%" & Chr(34) & " stop-color=" & Chr(34) & "rgba(0,0,0,0)" & Chr(34) & "/>" & Chr(13)
    SVG_Generator = SVG_Generator & "</linearGradient>" & Chr(13)
    SVG_Generator = SVG_Generator & "<linearGradient id=" & Chr(34) & "GS2" & Chr(34) & " x1=" & Chr(34) & "0" & Chr(34) & " x2=" & Chr(34) & "0" & Chr(34) & " y1=" & Chr(34) & "300000" & Chr(34) & " y2=" & Chr(34) & "100000" & Chr(34) & " gradientUnits=" & Chr(34) & "userSpaceOnUse" & Chr(34) & ">" & Chr(13)
    SVG_Generator = SVG_Generator & "<stop offset=" & Chr(34) & "0%" & Chr(34) & " stop-color=" & Chr(34) & "rgba(0,0,0,.1)" & Chr(34) & "/>" & Chr(13)
    SVG_Generator = SVG_Generator & "<stop offset=" & Chr(34) & "100%" & Chr(34) & " stop-color=" & Chr(34) & "rgba(0,0,0,0)" & Chr(34) & "/>" & Chr(13)
    SVG_Generator = SVG_Generator & "</linearGradient>" & Chr(13)
    
    SVG_Generator = SVG_Generator & "</defs>" & Chr(13)
    
    SVG_Generator = SVG_Generator & "<path id=" & Chr(34) & "theMotionPath" & Chr(34) & " d=" & Chr(34) & "M0 0H-" & ActivePlotter.PlotterWorkWidth & "H0" & Chr(34) & " />" & Chr(13)
    
    'Groups
    GNumber = 1
    SVG_Generator = SVG_Generator & SVG_group(ActivePlotter.Curves, GNumber)
    
    AnimationTxt = "<animateMotion dur=" & Chr(34) & "5s" & Chr(34) & " repeatCount=" & Chr(34) & "indefinite" & Chr(34) & ">" & Chr(13)
    AnimationTxt = AnimationTxt & "<mpath xlink:href=" & Chr(34) & "#theMotionPath" & Chr(34) & " />" & Chr(13)
    AnimationTxt = AnimationTxt & "</animateMotion>" & Chr(13)
    SVG_Generator = SVG_Generator & SVG_group(ActivePlotter.CarriageCurves, GNumber, AnimationTxt)
    
    'Material
    Dim R As Long, g As Long, b As Long
    Randomize Timer
    R = Int(Rnd * 255)
    g = Int(Rnd * 255)
    b = Int(Rnd * 255)
    SVG_Generator = SVG_Generator & "<rect fill=" & Chr(34) & "rgba(" & R & "," & g & "," & b & ",1)" & Chr(34) & " x=" & Chr(34) & -ActivePlotter.PlotterWorkWidth - 25000 & Chr(34) & " y=" & Chr(34) & "0" & Chr(34) & " width=" & Chr(34) & ActivePlotter.PlotterWorkWidth + 50000 & Chr(34) & " height=" & Chr(34) & "300000" & Chr(34) & "/>" & Chr(13)
    'Logo
    GNumber = 1
    
    Dim Brightness As Long
    Brightness = (0.2126 * R + 0.7152 * g + 0.0722 * b)
    If Brightness < 160 Then
        'Dark Material
        R = 255
        g = 255
        b = 255
    Else
        'Light Material
        R = 0
        g = 0
        b = 0
    End If
    
    Dim tsm As String
    Dim dopTsm As String
    Dim SpecialFill As String
    If ActivePlotter.PlotterWorkWidth < 600000 Then
        dopTsm = " scale(0.5)"
    End If
    tsm = " transform=" & Chr(34) & "translate(" & -ActivePlotter.PlotterWorkWidth / 2 & " 150000)" & dopTsm & Chr(34)
    SpecialFill = "stroke-width=" & Chr(34) & "1000" & Chr(34) & " stroke=" & Chr(34) & "rgba(" & R & "," & g & "," & b & ",1)" & Chr(34) & " fill=" & Chr(34) & "none" & Chr(34)
    SVG_Generator = SVG_Generator & SVG_group(GetLogoCurves, GNumber, , tsm, SpecialFill)
    
    SVG_Generator = SVG_Generator & "<rect fill=" & Chr(34) & "url(#GS1)" & Chr(34) & " x=" & Chr(34) & -ActivePlotter.PlotterWorkWidth - 25000 & Chr(34) & " y=" & Chr(34) & "0" & Chr(34) & " width=" & Chr(34) & ActivePlotter.PlotterWorkWidth + 50000 & Chr(34) & " height=" & Chr(34) & "25000" & Chr(34) & " />" & Chr(13)
    SVG_Generator = SVG_Generator & "<rect fill=" & Chr(34) & "url(#GS2)" & Chr(34) & " x=" & Chr(34) & -ActivePlotter.PlotterWorkWidth - 25000 & Chr(34) & " y=" & Chr(34) & "100000" & Chr(34) & " width=" & Chr(34) & ActivePlotter.PlotterWorkWidth + 50000 & Chr(34) & " height=" & Chr(34) & "200000" & Chr(34) & " />" & Chr(13)
    
    'Axis
    GNumber = 1

    If lZeroPosition = RightP Then
       tsm = " transform=" & Chr(34) & "translate(0 275000) rotate(-90) scale(3)" & Chr(34)
    ElseIf lZeroPosition = CenterP Then
       tsm = " transform=" & Chr(34) & "translate(" & -ActivePlotter.PlotterWorkWidth / 2 & " 150000) scale(3)" & Chr(34)
    Else
       tsm = " transform=" & Chr(34) & "translate(" & -ActivePlotter.PlotterWorkWidth & " 275000) scale(3)" & Chr(34)
    End If

    SVG_Generator = SVG_Generator & SVG_group(AxisCurves, GNumber, , tsm)
    SVG_Generator = SVG_Generator & "</svg>"

End Property
