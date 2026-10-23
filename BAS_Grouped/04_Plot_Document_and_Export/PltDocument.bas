' =============================================================================
' Group: 04_Plot_Document_and_Export
' Source module: PltDocument.cls
' Procedures: 71
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 84-98
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    HorisontalDelitel = 2
    DelayMultiplier = 1
    lHWndForm = GetActiveWindow
    hDC = GetDC(lHWndForm)
    Set myGraphics = New waGraphics
    Set AxisCurves = GetAxisCurves
    lViewAxis = True
    Set lBackgroundBrush = CreateBrush(ColorFromOle(SystemColorConstants.vbScrollBars))
    Set MaterialBrush = CreateBrush(RGB(255, 255, 255))
    
    lShadowDistance = 5
    lViewBorder = True
    lBorderColor = ColorFromOle(SystemColorConstants.vbWindowFrame)
End Sub

' -----------------------------------------------------------------------------
' SVG_Generator (PropertyGet), source lines 102-191
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' SVG_Gradient (Function), source lines 192-207
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' SVG_group (Function), source lines 209-302
' -----------------------------------------------------------------------------
Private Function SVG_group(Curves As Collection, ByRef GNumber As Long, Optional AnimationTxt As String = "", Optional Transform As String = "", Optional SpecilFill As String = "") As String
    Dim curve As waPath
    Dim segments() As segment
    Dim q As Long
    Dim SVGtxt As String
    Dim lastType As SegmentType
    Dim LastSVG_X As Long
    Dim LastSVG_Y As Long
    Dim startX As Long
    Dim startY As Long

    Dim LastLetter As String
    Dim Color As String
    
                
    SVGtxt = SVGtxt & "<g" & Transform & ">" & Chr(13)
    For Each curve In Curves
        LastSVG_X = 0
        LastSVG_Y = 0
        LastLetter = ""
        With curve
            If .brush.FillType = BrushTypeSolidColor Then
                Color = rgbaToStr(.brush.Color)
            ElseIf .brush.FillType = BrushTypeLinearGradient Then
                Color = "url(#G" & GNumber & ")"
                GNumber = GNumber + 1
            Else
                Color = "rgba(0,0,0,1)"
            End If
            If SpecilFill = "" Then
                SVGtxt = SVGtxt & "<path fill=" & Chr(34) & Color & Chr(34) & " d=" & Chr(34)
            Else
                SVGtxt = SVGtxt & "<path " & SpecilFill & " d=" & Chr(34)
            End If
        
            segments = .segments
            For q = 0 To UBound(segments)
                With segments(q)
                    Select Case .Type
                        
                        Case InvisibleLine
                            SVGtxt = SVGtxt & "m" & .P2.x - LastSVG_X
                            SVGtxt = SVGtxt & AddValue(-.P2.y - LastSVG_Y, SVGtxt)
                            LastSVG_X = .P2.x
                            LastSVG_Y = -.P2.y
                            startX = LastSVG_X
                            startY = LastSVG_Y
                        Case LineSegment
                            'Is need close curve?
                            If NeedCloseCurve(segments(), q, startX, startY) Then
                                SVGtxt = SVGtxt & "z"
                            Else
                                'hv
                                If LastSVG_X = .P2.x Then
                                    SVGtxt = SVGtxt & "v" & -.P2.y - LastSVG_Y
                                    LastLetter = "v"
                                ElseIf -LastSVG_Y = .P2.y Then
                                    SVGtxt = SVGtxt & "h" & .P2.x - LastSVG_X
                                    LastLetter = "h"
                                Else
                                    'else l
                                    If lastType <> .Type Or LastLetter <> "l" Then SVGtxt = SVGtxt & "l"
                                    SVGtxt = SVGtxt & AddValue(.P2.x - LastSVG_X, SVGtxt)
                                    SVGtxt = SVGtxt & AddValue(-.P2.y - LastSVG_Y, SVGtxt)
                                    LastLetter = "l"
                                End If
                            End If
                            

                            LastSVG_X = .P2.x
                            LastSVG_Y = -.P2.y
                            
                        Case CurveSegment
                            If lastType <> .Type Then SVGtxt = SVGtxt & "c"
                            SVGtxt = SVGtxt & AddValue(.P2.x - LastSVG_X, SVGtxt)
                            SVGtxt = SVGtxt & AddValue(-.P2.y - LastSVG_Y, SVGtxt)
                            SVGtxt = SVGtxt & AddValue(.P3.x - LastSVG_X, SVGtxt)
                            SVGtxt = SVGtxt & AddValue(-.P3.y - LastSVG_Y, SVGtxt)
                            SVGtxt = SVGtxt & AddValue(.P4.x - LastSVG_X, SVGtxt)
                            SVGtxt = SVGtxt & AddValue(-.P4.y - LastSVG_Y, SVGtxt)
                            LastSVG_X = .P4.x
                            LastSVG_Y = -.P4.y
                            
                    End Select
                    lastType = .Type
                End With
            Next
        End With
        SVGtxt = SVGtxt & Chr(34) & "/>" & Chr(13)
    Next
    SVGtxt = SVGtxt & AnimationTxt
    SVGtxt = SVGtxt & "</g>" & Chr(13)
    SVG_group = SVGtxt
End Function

' -----------------------------------------------------------------------------
' NeedCloseCurve (Function), source lines 303-309
' -----------------------------------------------------------------------------
Private Function NeedCloseCurve(ByRef segments() As segment, q As Long, startX As Long, startY As Long) As Boolean
    If q = UBound(segments()) Then
        If segments(q).P2.x = startX And -segments(q).P2.y = startY Then NeedCloseCurve = True
    Else
        If segments(q + 1).Type = InvisibleLine And segments(q).P2.x = startX And -segments(q).P2.y = startY Then NeedCloseCurve = True
    End If
End Function

' -----------------------------------------------------------------------------
' rgbaToStr (Function), source lines 310-320
' -----------------------------------------------------------------------------
Private Function rgbaToStr(Color As RGBA) As String
    Dim s As Double
    Dim aa As String
    
    With Color
        s = .a / 255
        aa = Replace(Format(s, "#0.00"), ",", ".")
        If aa = "1.00" Then aa = 1
        rgbaToStr = "rgba(" & .R & "," & .g & "," & .b & "," & aa & ")"
    End With
End Function

' -----------------------------------------------------------------------------
' AddValue (Function), source lines 321-334
' -----------------------------------------------------------------------------
Private Function AddValue(value As Long, ByRef SVGtxt As String) As String
    Dim rs As Long
    rs = Asc(right(SVGtxt, 1))
    If rs > 47 And rs < 58 Then
        'need separator
        If value < 0 Then
            AddValue = value
        Else
            AddValue = " " & value
        End If
    Else
        AddValue = value
    End If
End Function

' -----------------------------------------------------------------------------
' StopShowWay (Sub), source lines 336-338
' -----------------------------------------------------------------------------
Public Sub StopShowWay()
    ShowWayProcessing = False
End Sub

' -----------------------------------------------------------------------------
' ShowWay (Sub), source lines 339-416
' -----------------------------------------------------------------------------
Public Sub ShowWay()
    If ShowWayProcessing Then
        ShowWayProcessing = False
        ShowWayShowNow = False
        Render
        Exit Sub
    End If
    
    If ShowWayShowNow Then Render
    ShowWayProcessing = True
    ShowWayShowNow = False
    Dim curvesMatrix As New waMatrix
    Dim lastP As POINTL
    
    With curvesMatrix
        .ScaleM 1 / mScale, -1 / mScale
        Select Case lZeroPosition
            Case RightP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale - lYmargin / mScale, OnScreenOrignY
                lastP.x = AllCurvesWidth + lYmargin
                lastP.y = -lXmargin
            Case LeftP
                .Translate OnScreenOrignX + lXmargin / mScale, OnScreenOrignY
                lastP.x = -lXmargin
                lastP.y = -lYmargin
            Case CenterP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale / 2 + lXmargin / mScale, OnScreenOrignY + AllCurvesHeight / mScale / 2
                lastP.x = -lXmargin + AllCurvesWidth / 2
                lastP.y = -lYmargin + AllCurvesHeight / 2
        End Select
    End With
    myGraphics.SetWorldTransform curvesMatrix
        
    Dim segments() As segment
    Dim path As waPath
    Dim q1 As Long
    Dim q2 As Long
    Dim q As Long
    
    For q1 = 1 To Curves.count
        Set path = Curves(q1)
        segments() = path.segments()
        segments(0).P1 = lastP
        For q2 = 0 To UBound(segments)
            With segments(q2)
                Select Case .Type
                    Case InvisibleLine
                        myGraphics.DrawLine UpWayPen, .P1.x, .P1.y, .P2.x, .P2.y
                    Case LineSegment
                        myGraphics.DrawLine WayPen, .P1.x, .P1.y, .P2.x, .P2.y
                    Case CurveSegment
                        myGraphics.DrawBezier WayPen, .P1.x, .P1.y, .P2.x, .P2.y, .P3.x, .P3.y, .P4.x, .P4.y
                End Select
            End With
            Refresh
            For q = 1 To DelayMultiplier
                DoEvents
                If Not ShowWayProcessing Then
                    If DelayMultiplier = 20 Then DelayMultiplier = 1 Else DelayMultiplier = 20
                    Exit Sub
                End If
                Sleep 10
            Next
        Next
        'Last Point for line or bezier
        With segments(UBound(segments))
            If .Type = CurveSegment Then
                lastP = .P4
            Else
                lastP = .P2
            End If
        End With
    Next
    ShowWayProcessing = False
    ShowWayShowNow = True
    If DelayMultiplier = 20 Then DelayMultiplier = 1 Else DelayMultiplier = 20
    
End Sub

' -----------------------------------------------------------------------------
' SetCurvesByVal (Sub), source lines 417-464
' -----------------------------------------------------------------------------
Public Sub SetCurvesByVal(OriginalCurves As Collection, Height As Long, Width As Long, Optional Turn As Turns, Optional VerticalMirror As Boolean, Optional HorizontalMirror As Boolean)

    Set Curves = CloneCollection(OriginalCurves)
    
    If Turn <> Turn0 Or VerticalMirror Or HorizontalMirror Then
        'Need create Matrix
        Dim mx As New waMatrix
    End If
    
    If Turn = Turn0 Or Turn = Turn180 Then
        AllCurvesHeight = Height
        AllCurvesWidth = Width
    Else
        AllCurvesHeight = Width
        AllCurvesWidth = Height
    End If
    
    'Turn
    If Turn = TurnMinus90 Then
        mx.Rotate 90
        mx.Translate AllCurvesWidth, 0
    ElseIf Turn = Turn180 Then
        mx.Rotate 180
        mx.Translate AllCurvesWidth, AllCurvesHeight
    ElseIf Turn = TurnPlus90 Then
        mx.Rotate 270
        mx.Translate 0, AllCurvesHeight
    End If
    
    'Mirror
    If HorizontalMirror Then
        mx.ScaleM -1, 1
        mx.Translate AllCurvesWidth, 0
    End If
    
    If VerticalMirror Then
        mx.ScaleM 1, -1
        mx.Translate 0, AllCurvesHeight
    End If
    
    If Turn <> Turn0 Or VerticalMirror Or HorizontalMirror Then
        'Need apply Matrix
        Dim path As waPath
        For Each path In Curves
            path.ApplyMatrixTransform mx
        Next
    End If
End Sub

' -----------------------------------------------------------------------------
' CloneCollection (Function), source lines 465-473
' -----------------------------------------------------------------------------
Private Function CloneCollection(col As Collection) As Collection
    Set CloneCollection = New Collection
    Dim path As waPath
    Dim newPath As waPath
    For Each path In col
        Set newPath = path.Clone
        CloneCollection.Add newPath
    Next
End Function

' -----------------------------------------------------------------------------
' MaterialBrush (PropertySet), source lines 474-502
' -----------------------------------------------------------------------------
Property Set MaterialBrush(oValue As waBrush)
    Set lMaterialBrush = oValue
    
    Dim b As Long
    b = lMaterialBrush.Brightness
    
    If b < 150 Then
        'Dark Material
        Set CurvesPen = CreatePen(RGB(210, 210, 210), 80, 1.6)
        Set WayPen = CreatePen(RGB(255, 237, 0), 50, 2)
        Set UpWayPen = CreatePen(RGB(255, 237, 0), 25, 2)
        
    Else
        'Light Material
        Set CurvesPen = CreatePen(RGB(110, 110, 110), 80, 1.6)
        Set WayPen = CreatePen(RGB(57, 49, 133), 50, 2)
        Set UpWayPen = CreatePen(RGB(57, 49, 133), 30, 2)
    End If
    
    UpWayPen.DashStyle = Dash
    UpWayPen.EndCap = ArrowAnchor
    WayPen.EndCap = ArrowAnchor
    
    Dim m As New waMatrix
    m.ScaleM 0.025, 0.025
    Set WayPen.matrix = m
    Set UpWayPen.matrix = m
        
End Property

' -----------------------------------------------------------------------------
' MaterialBrush (PropertyGet), source lines 503-505
' -----------------------------------------------------------------------------
Property Get MaterialBrush() As waBrush
    Set MaterialBrush = lMaterialBrush
End Property

' -----------------------------------------------------------------------------
' SetRandomMaterialColor (Sub), source lines 506-518
' -----------------------------------------------------------------------------
Public Sub SetRandomMaterialColor()
    Randomize Timer
    Dim R As Byte, g As Byte, b As Byte, a As Long
    If Int(Rnd * 10) = 1 Then
        a = Int(Rnd * 50) + 50
    Else
        a = 100
    End If
    R = Int(Rnd * 255)
    g = Int(Rnd * 255)
    b = Int(Rnd * 255)
    Set MaterialBrush = CreateBrush(RGB(R, g, b), a)
End Sub

' -----------------------------------------------------------------------------
' BorderColor (PropertyLet), source lines 520-522
' -----------------------------------------------------------------------------
Property Let BorderColor(oValue As Long)
    lBorderColor = oValue
End Property

' -----------------------------------------------------------------------------
' BorderColor (PropertyGet), source lines 523-525
' -----------------------------------------------------------------------------
Property Get BorderColor() As Long
    BorderColor = lBorderColor
End Property

' -----------------------------------------------------------------------------
' BackgroundBrush (PropertySet), source lines 526-528
' -----------------------------------------------------------------------------
Property Set BackgroundBrush(oValue As waBrush)
    Set lBackgroundBrush = oValue
End Property

' -----------------------------------------------------------------------------
' BackgroundBrush (PropertyGet), source lines 529-531
' -----------------------------------------------------------------------------
Property Get BackgroundBrush() As waBrush
    Set BackgroundBrush = lBackgroundBrush
End Property

' -----------------------------------------------------------------------------
' ViewBorder (PropertyLet), source lines 532-534
' -----------------------------------------------------------------------------
Property Let ViewBorder(oValue As Boolean)
    lViewBorder = oValue
End Property

' -----------------------------------------------------------------------------
' ViewBorder (PropertyGet), source lines 535-537
' -----------------------------------------------------------------------------
Property Get ViewBorder() As Boolean
    ViewBorder = lViewBorder
End Property

' -----------------------------------------------------------------------------
' ShadowDistance (PropertyLet), source lines 538-540
' -----------------------------------------------------------------------------
Property Let ShadowDistance(oValue As Long)
    lShadowDistance = oValue
End Property

' -----------------------------------------------------------------------------
' ShadowDistance (PropertyGet), source lines 541-543
' -----------------------------------------------------------------------------
Property Get ShadowDistance() As Long
    ShadowDistance = lShadowDistance
End Property

' -----------------------------------------------------------------------------
' ViewAxis (PropertyGet), source lines 544-546
' -----------------------------------------------------------------------------
Property Get ViewAxis() As Boolean
    ViewAxis = lViewAxis
End Property

' -----------------------------------------------------------------------------
' ViewAxis (PropertyLet), source lines 547-549
' -----------------------------------------------------------------------------
Property Let ViewAxis(oValue As Boolean)
    lViewAxis = oValue
End Property

' -----------------------------------------------------------------------------
' ViewLogo (PropertyGet), source lines 550-552
' -----------------------------------------------------------------------------
Property Get ViewLogo() As Boolean
    ViewLogo = lViewLogo
End Property

' -----------------------------------------------------------------------------
' ViewLogo (PropertyLet), source lines 553-555
' -----------------------------------------------------------------------------
Property Let ViewLogo(oValue As Boolean)
    lViewLogo = oValue
End Property

' -----------------------------------------------------------------------------
' RedLineError (PropertyGet), source lines 556-558
' -----------------------------------------------------------------------------
Property Get RedLineError() As Boolean
    RedLineError = lRedLineError
End Property

' -----------------------------------------------------------------------------
' Xmargin (PropertyGet), source lines 559-561
' -----------------------------------------------------------------------------
Property Get Xmargin() As Long
    Xmargin = lXmargin
End Property

' -----------------------------------------------------------------------------
' Xmargin (PropertyLet), source lines 562-569
' -----------------------------------------------------------------------------
Property Let Xmargin(oValue As Long)
    If lXmargin <> oValue Then
        lXmargin = oValue
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property

' -----------------------------------------------------------------------------
' Ymargin (PropertyGet), source lines 570-572
' -----------------------------------------------------------------------------
Property Get Ymargin() As Long
    Ymargin = lYmargin
End Property

' -----------------------------------------------------------------------------
' Ymargin (PropertyLet), source lines 573-580
' -----------------------------------------------------------------------------
Property Let Ymargin(oValue As Long)
    If lYmargin <> oValue Then
        lYmargin = oValue
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property

' -----------------------------------------------------------------------------
' XYmargin (PropertyGet), source lines 581-584
' -----------------------------------------------------------------------------
Property Get XYmargin() As POINTL
    XYmargin.x = lXmargin
    XYmargin.y = lYmargin
End Property

' -----------------------------------------------------------------------------
' XYmargin (PropertyLet), source lines 585-593
' -----------------------------------------------------------------------------
Property Let XYmargin(oValue As POINTL)
    If lXmargin <> oValue.x Or lYmargin <> oValue.y Then
        lXmargin = oValue.x
        lYmargin = oValue.y
        MarginCheck
        Render
        ShowWayProcessing = False
    End If
End Property

' -----------------------------------------------------------------------------
' MarginCheck (Sub), source lines 594-612
' -----------------------------------------------------------------------------
Private Sub MarginCheck()
    Select Case lZeroPosition
        Case RightP
            If lXmargin < 0 Then lXmargin = 0
            If lYmargin < 0 Then lYmargin = 0
            If lXmargin > MaxSheetHeight Then lXmargin = MaxSheetHeight
            If lYmargin > MaxSheetWidth Then lYmargin = MaxSheetWidth
        Case LeftP
            If lXmargin < 0 Then lXmargin = 0
            If lYmargin < 0 Then lYmargin = 0
            If lXmargin > MaxSheetWidth Then lXmargin = MaxSheetWidth
            If lYmargin > MaxSheetHeight Then lYmargin = MaxSheetHeight
        Case CenterP
            If lXmargin < -MaxSheetWidth / 2 Then lXmargin = -MaxSheetWidth / 2
            If lYmargin < -MaxSheetHeight / 2 Then lYmargin = -MaxSheetHeight / 2
            If lXmargin > MaxSheetWidth / 2 Then lXmargin = MaxSheetWidth / 2
            If lYmargin > MaxSheetHeight / 2 Then lYmargin = MaxSheetHeight / 2
    End Select
End Sub

' -----------------------------------------------------------------------------
' ZeroPosition (PropertyLet), source lines 613-630
' -----------------------------------------------------------------------------
Property Let ZeroPosition(oValue As ZeroPositions)
    lZeroPosition = oValue
    
    'Init Plotter
    If PlotterModelString = "" Then
        If MaxSheetWidth > 600000 Then
            ActivePlotter = GetModel_1600
        Else
            ActivePlotter = GetModel_600
        End If
    Else
        ActivePlotter = GetPlotterFromString(PlotterModelString, PlotterName)
        ActivePlotter.PlotterWorkWidth = MaxSheetWidth
    End If
    
    CalculatePage
    Sorting
End Property

' -----------------------------------------------------------------------------
' ZeroPosition (PropertyGet), source lines 631-633
' -----------------------------------------------------------------------------
Property Get ZeroPosition() As ZeroPositions
    ZeroPosition = lZeroPosition
End Property

' -----------------------------------------------------------------------------
' SetScreenSize (Sub), source lines 634-639
' -----------------------------------------------------------------------------
Public Sub SetScreenSize(ScreenWidth As Long, ScreenHeight As Long)
    Me.ScreenWidth = ScreenWidth
    Me.ScreenHeight = ScreenHeight
    MemoryBitmap1.SetSize ScreenWidth, ScreenHeight
    myGraphics.hDC = MemoryBitmap1.hDC
End Sub

' -----------------------------------------------------------------------------
' CalculatePage (Sub), source lines 640-709
' -----------------------------------------------------------------------------
Private Sub CalculatePage()
    Dim rScreenW As Long, rScreenH As Long
    rScreenW = ScreenWidth - FromFrameOffset - FromFrameOffset
    rScreenH = ScreenHeight - FromFrameOffset - FromFrameOffset
    
    If Curves.count = 0 Then
        If ActivePlotter.Height / rScreenH > ActivePlotter.Width / rScreenW Then
            'Vertical
            mScale = CDbl(ActivePlotter.Height) / rScreenH
            If lViewLogo Then
                OnScreenSheetHeight = (ActivePlotter.Height * 0.3 + 2 * WhiteOffset) / mScale
            Else
                OnScreenSheetHeight = WhiteOffset / mScale
            End If
            OnScreenSheetWidth = (MaxSheetWidth + WhiteOffset + WhiteOffset) / mScale
                       
            OnScreenSheetPositionX = ScreenWidth / 2 + (Abs(ActivePlotter.left) - MaxSheetWidth - WhiteOffset - ActivePlotter.Width / 2) / mScale
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        Else
            'Horisontal
            mScale = CDbl(ActivePlotter.Width) / rScreenW
            If lViewLogo Then
                OnScreenSheetHeight = (ActivePlotter.Height * 0.3 + 2 * WhiteOffset) / mScale
            Else
                OnScreenSheetHeight = WhiteOffset / mScale
            End If
            OnScreenSheetWidth = (MaxSheetWidth + WhiteOffset + WhiteOffset) / mScale
            
            OnScreenSheetPositionX = FromFrameOffset + (Abs(ActivePlotter.left) - MaxSheetWidth - WhiteOffset) / mScale
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        End If
    Else
        Dim sheetHeight As Long
        If AllCurvesHeight > MaxSheetHeight Then
            sheetHeight = MaxSheetHeight
        Else
            sheetHeight = AllCurvesHeight
        End If
    
        If (sheetHeight + ActivePlotter.top + 2 * WhiteOffset) / rScreenH > (MaxSheetWidth + 2 * WhiteOffset) / rScreenW Then
            'Vertical
            mScale = CDbl(sheetHeight + ActivePlotter.top + 2 * WhiteOffset) / rScreenH
            OnScreenSheetHeight = rScreenH - ActivePlotter.top / mScale
            OnScreenSheetWidth = (MaxSheetWidth + 2 * WhiteOffset) / mScale
            
            OnScreenSheetPositionX = ScreenWidth / 2 - OnScreenSheetWidth / 2
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        Else
            'Horisontal
            mScale = CDbl(MaxSheetWidth + 2 * WhiteOffset) / rScreenW
            OnScreenSheetHeight = (sheetHeight + 2 * WhiteOffset) / mScale
            OnScreenSheetWidth = rScreenW
            
            OnScreenSheetPositionX = FromFrameOffset
            OnScreenSheetPositionY = FromFrameOffset + ActivePlotter.top / mScale
        End If
    End If
    
    Select Case lZeroPosition
        Case RightP
            OnScreenOrignX = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale
        Case LeftP
            OnScreenOrignX = OnScreenSheetPositionX + WhiteOffset / mScale
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale
        Case CenterP
            OnScreenOrignX = OnScreenSheetPositionX + OnScreenSheetWidth / 2
            OnScreenOrignY = OnScreenSheetPositionY + OnScreenSheetHeight / 2
    End Select
End Sub

' -----------------------------------------------------------------------------
' Render (Sub), source lines 710-726
' -----------------------------------------------------------------------------
Public Sub Render()
    Static t As Single
    t = Timer
    DrawScreen
    DrawPlotterShadow
    DrawSheetShadow
    DrawPlotter
    DrawCarriage
    DrawSheet
    DrawCurves
    DrawLogo
    DrawRedLines
    DrawTopMaterialShadow
    DrawAxis
    Refresh
    RaiseEvent RenderTime(Timer - t)
End Sub

' -----------------------------------------------------------------------------
' Refresh (Sub), source lines 727-729
' -----------------------------------------------------------------------------
Public Sub Refresh()
    BitBlt hDC, PositionX, PositionY, MemoryBitmap1.Width, MemoryBitmap1.Height, MemoryBitmap1.hDC, 0, 0, SRCCOPY
End Sub

' -----------------------------------------------------------------------------
' DrawScreen (Sub), source lines 730-735
' -----------------------------------------------------------------------------
Private Sub DrawScreen()
    myGraphics.ResetClip
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.FillRect lBackgroundBrush, -1, -1, ScreenWidth + 1, ScreenHeight + 1
    If lViewBorder Then myGraphics.DrawRect CreatePen(lBorderColor), 0, 0, ScreenWidth - 1, ScreenHeight - 1
End Sub

' -----------------------------------------------------------------------------
' DrawSheet (Sub), source lines 736-766
' -----------------------------------------------------------------------------
Private Sub DrawSheet()
    If Curves.count > 0 Or lViewLogo Then
        Dim hPlus As Long
        If lZeroPosition = RightP Then
            hPlus = lXmargin / mScale
        Else
            If lZeroPosition = CenterP Then
                If lYmargin < 0 Then
                    hPlus = 0
                Else
                    hPlus = lYmargin / mScale
                End If
            Else
                hPlus = lYmargin / mScale
            End If
        End If
    
        myGraphics.SetWorldTransform OnScreenMatrix
        myGraphics.SetClip 0, 0, ScreenWidth - 1, ScreenHeight - 1
        
        myGraphics.FillRect lMaterialBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
        
'        Dim gradientBrush As New waBrush
'        gradientBrush.SetLinearGradient OnScreenSheetPositionX, OnScreenSheetPositionY + WhiteOffset / mScale + 1, OnScreenSheetPositionX, OnScreenSheetPositionY, 0, 0, , 0, 30
        
'        myGraphics.SetClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + hPlus + 1
'        myGraphics.FillRect gradientBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, WhiteOffset / mScale
'        myGraphics.SetClip 0, 0, ScreenWidth - 1, ScreenHeight - 1
        myGraphics.DrawRect CreatePen(RGB(150, 150, 150), 100), OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawTopMaterialShadow (Sub), source lines 767-774
' -----------------------------------------------------------------------------
Private Sub DrawTopMaterialShadow()
    If Curves.count > 0 Or lViewLogo Then
        Dim gradientBrush As New waBrush
        gradientBrush.SetLinearGradient OnScreenSheetPositionX, OnScreenSheetPositionY + WhiteOffset / mScale + 1, OnScreenSheetPositionX, OnScreenSheetPositionY, 0, 0, , 0, 30
        myGraphics.SetClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, OnScreenSheetHeight
        myGraphics.FillRect gradientBrush, OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth, WhiteOffset / mScale
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawSheetShadow (Sub), source lines 775-796
' -----------------------------------------------------------------------------
Private Sub DrawSheetShadow()
    If Curves.count > 0 Or lViewLogo Then
        Dim hPlus As Long
        If lZeroPosition = RightP Then
            hPlus = lXmargin / mScale
        Else
            If lZeroPosition = CenterP Then
                If lYmargin < 0 Then
                    hPlus = 0
                Else
                    hPlus = lYmargin / mScale
                End If
            Else
                hPlus = lYmargin / mScale
            End If
        End If
        
        myGraphics.SetWorldTransform OnScreenMatrix
        myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
        myGraphics.FillRect CreateBrush(RGB(170, 170, 170), 100), OnScreenSheetPositionX + lShadowDistance, OnScreenSheetPositionY + lShadowDistance, OnScreenSheetWidth, OnScreenSheetHeight + hPlus
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawCurves (Sub), source lines 797-831
' -----------------------------------------------------------------------------
Private Sub DrawCurves()
    Dim curvesMatrix As New waMatrix
    
    With curvesMatrix
        .ScaleM 1 / mScale, -1 / mScale
        Select Case lZeroPosition
            Case RightP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale - lYmargin / mScale, OnScreenOrignY
            Case LeftP
                .Translate OnScreenOrignX + lXmargin / mScale, OnScreenOrignY
            Case CenterP
                .Translate OnScreenOrignX - AllCurvesWidth / mScale / 2 + lXmargin / mScale, OnScreenOrignY + AllCurvesHeight / mScale / 2
        End Select
    End With
    
    Dim hPlus As Long
    If lZeroPosition = RightP Then
        hPlus = lXmargin / mScale
    Else
        If lZeroPosition = CenterP Then
            If lYmargin < 0 Then
                hPlus = 0
            Else
                hPlus = lYmargin / mScale
            End If
        Else
            hPlus = lYmargin / mScale
        End If
    End If
    
    SetOnScreenClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + 1 + hPlus
    Set myGraphics.Curves = Curves
    myGraphics.SetWorldTransform curvesMatrix
    myGraphics.Render , , , CurvesPen
End Sub

' -----------------------------------------------------------------------------
' DrawLogo (Sub), source lines 832-865
' -----------------------------------------------------------------------------
Private Sub DrawLogo()
    If lViewLogo And Curves.count = 0 Then
        Dim k As Single
        Dim curvesMatrix As New waMatrix
        k = 700000 / (MaxSheetWidth + 2 * WhiteOffset)
        
        With curvesMatrix
            .ScaleM 1 / mScale / k, -1 / mScale / k
            .Translate OnScreenSheetPositionX + OnScreenSheetWidth / 2, OnScreenSheetPositionY + OnScreenSheetHeight / 2
        End With
        
        myGraphics.SetWorldTransform OnScreenMatrix
        SetOnScreenClip OnScreenSheetPositionX, OnScreenSheetPositionY, OnScreenSheetWidth + 1, OnScreenSheetHeight + 1
        Set myGraphics.Curves = GetLogoCurves
        
        Dim P1 As waPath
        Dim P2 As waPath
        Set P1 = myGraphics.Curves(1)
        Set P2 = myGraphics.Curves(2)
        
        Dim c As Long
        c = CurvesPen.ColorARGB
        
        Set P1.pen = CreatePen(, 100, 1.6)
        Set P2.pen = CreatePen(, 100, 2)
        
        P1.pen.ColorARGB = c
        P2.pen.ColorARGB = c
        
        
        myGraphics.SetWorldTransform curvesMatrix
        myGraphics.Render
    End If
End Sub

' -----------------------------------------------------------------------------
' DrawRedLines (Sub), source lines 866-983
' -----------------------------------------------------------------------------
Private Sub DrawRedLines()
    Dim LV As Boolean
    Dim RV As Boolean
    Dim UH As Boolean
    Dim BH As Boolean
    
    Dim dy As Long
    Dim dY2 As Long
    Dim margin As Long
    
    Dim sheetHeight As Long
    If AllCurvesHeight > MaxSheetHeight Then
        sheetHeight = MaxSheetHeight
    Else
        sheetHeight = AllCurvesHeight
    End If
    
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 1, ScreenHeight - 1
    
    Select Case lZeroPosition
        Case RightP
            If AllCurvesWidth + lYmargin > MaxSheetWidth Then LV = True
            If AllCurvesHeight + lXmargin > MaxSheetHeight Then UH = True
            dy = (lXmargin - (MaxSheetHeight - sheetHeight)) / mScale
            margin = lXmargin / mScale
        Case LeftP
            If AllCurvesWidth + lXmargin > MaxSheetWidth Then RV = True
            If AllCurvesHeight + lYmargin > MaxSheetHeight Then UH = True
            dy = (lYmargin - (MaxSheetHeight - sheetHeight)) / mScale
            margin = lYmargin / mScale
        Case CenterP
            If AllCurvesWidth / 2 + lXmargin > MaxSheetWidth / 2 Then RV = True
            If -AllCurvesWidth / 2 + lXmargin < -MaxSheetWidth / 2 Then LV = True
            
            If AllCurvesHeight / 2 + lYmargin > MaxSheetHeight / 2 Then UH = True
            If -AllCurvesHeight / 2 + lYmargin < -MaxSheetHeight / 2 Then BH = True
            
            dy = (lYmargin - (MaxSheetHeight - sheetHeight) / 2) / mScale
            dY2 = (lYmargin + (MaxSheetHeight - sheetHeight) / 2) / mScale
            margin = lYmargin / mScale
            If margin < 0 Then margin = 0
    End Select

    Dim redPen As waPen, whitePen As waPen
    Set redPen = CreatePen(RGB(255, 0, 0), 90, 2)
    Set whitePen = CreatePen(RGB(255, 255, 255), 50, 6)
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    If LV Then
        x1 = OnScreenSheetPositionX + WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    
    If RV Then
        x1 = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    
    If UH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + WhiteOffset / mScale + dy
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        If y1 > OnScreenSheetPositionY Then
            myGraphics.DrawLine whitePen, x1, y1, x2, y2
        End If
    End If
    
    If BH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale + dY2
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        myGraphics.DrawLine whitePen, x1, y1, x2, y2
    End If
    'Red===================================================
    If LV Then
        x1 = OnScreenSheetPositionX + WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    
    If RV Then
        x1 = OnScreenSheetPositionX + OnScreenSheetWidth - WhiteOffset / mScale
        y1 = OnScreenSheetPositionY + 1
        x2 = x1
        y2 = OnScreenSheetPositionY + OnScreenSheetHeight + margin
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    
    If UH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + WhiteOffset / mScale + dy
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        If y1 > OnScreenSheetPositionY Then
            myGraphics.DrawLine redPen, x1, y1, x2, y2
        End If
    End If
    
    If BH Then
        x1 = OnScreenSheetPositionX
        y1 = OnScreenSheetPositionY + OnScreenSheetHeight - WhiteOffset / mScale + dY2
        x2 = OnScreenSheetPositionX + OnScreenSheetWidth
        y2 = y1
        myGraphics.DrawLine redPen, x1, y1, x2, y2
    End If
    lRedLineError = LV Or RV Or UH Or BH
End Sub

' -----------------------------------------------------------------------------
' DrawCarriage (Sub), source lines 984-1008
' -----------------------------------------------------------------------------
Private Sub DrawCarriage()
    'PlotterCarriageCurves
    Dim plotterMatrix As New waMatrix

    Set myGraphics.Curves = ActivePlotter.CarriageCurves
    
    Select Case lZeroPosition
        Case RightP
            plotterMatrix.Translate -WhiteOffset - lYmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth, OnScreenSheetPositionY
        Case LeftP
            plotterMatrix.Translate WhiteOffset + lXmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX, OnScreenSheetPositionY
        Case CenterP
            plotterMatrix.Translate lXmargin, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth \ 2, OnScreenSheetPositionY
    End Select
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
    myGraphics.SetWorldTransform plotterMatrix
    myGraphics.Render False, True
End Sub

' -----------------------------------------------------------------------------
' DrawPlotter (Sub), source lines 1009-1033
' -----------------------------------------------------------------------------
Private Sub DrawPlotter()
    Dim plotterMatrix As New waMatrix

    Set myGraphics.Curves = ActivePlotter.Curves
    
    
    Select Case lZeroPosition
        Case RightP
            plotterMatrix.Translate -WhiteOffset, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth, OnScreenSheetPositionY
        Case LeftP
            plotterMatrix.Translate WhiteOffset + ActivePlotter.PlotterWorkWidth, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX, OnScreenSheetPositionY
        Case CenterP
            plotterMatrix.Translate ActivePlotter.PlotterWorkWidth \ 2, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth \ 2, OnScreenSheetPositionY
    End Select
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
    myGraphics.SetWorldTransform plotterMatrix
    myGraphics.Render False, True
End Sub

' -----------------------------------------------------------------------------
' DrawPlotterShadow (Sub), source lines 1034-1056
' -----------------------------------------------------------------------------
Private Sub DrawPlotterShadow()
    Dim plotterMatrix As New waMatrix
    Set myGraphics.Curves = ActivePlotter.Curves

    Select Case lZeroPosition
        Case RightP
            plotterMatrix.Translate -WhiteOffset, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
        Case LeftP
            plotterMatrix.Translate WhiteOffset + ActivePlotter.PlotterWorkWidth, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
        Case CenterP
            plotterMatrix.Translate ActivePlotter.PlotterWorkWidth \ 2, 0
            plotterMatrix.ScaleM 1 / mScale, -1 / mScale
            plotterMatrix.Translate OnScreenSheetPositionX + OnScreenSheetWidth \ 2 + lShadowDistance, OnScreenSheetPositionY + lShadowDistance
    End Select
    myGraphics.SetWorldTransform OnScreenMatrix
    myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
    myGraphics.SetWorldTransform plotterMatrix
    myGraphics.Render , , CreateBrush(RGB(170, 170, 170), 100)
End Sub

' -----------------------------------------------------------------------------
' DrawAxis (Sub), source lines 1057-1079
' -----------------------------------------------------------------------------
Private Sub DrawAxis()
    If lViewAxis Then
        Dim hPlus As Long
        If lZeroPosition = RightP Then
            hPlus = lXmargin / mScale
        Else
            hPlus = lYmargin / mScale
        End If
            
        myGraphics.SetWorldTransform OnScreenMatrix
        myGraphics.SetClip 1, 1, ScreenWidth - 2, ScreenHeight - 2
        
        Set myGraphics.Curves = AxisCurves
        Dim myMatrix As New waMatrix
        myMatrix.ScaleM 0.002, -0.002
        
        If lZeroPosition = RightP Then myMatrix.Rotate -90
        myMatrix.Translate OnScreenOrignX, OnScreenOrignY + hPlus
        myGraphics.SetWorldTransform myMatrix
        
        myGraphics.Render False, True
    End If
End Sub

' -----------------------------------------------------------------------------
' SetOnScreenClip (Sub), source lines 1080-1093
' -----------------------------------------------------------------------------
Private Sub SetOnScreenClip(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
    If x < 1 Then
        Width = Width - (1 - x)
        x = 1
    End If
    If y < 1 Then
        Height = Height - (1 - y)
        y = 1
    End If

    If x + Width > ScreenWidth Then Width = ScreenWidth - x
    If y + Height > ScreenHeight Then Height = ScreenHeight - y
    myGraphics.SetClip x, y, Width, Height
End Sub

' -----------------------------------------------------------------------------
' Sorting (Sub), source lines 1094-1156
' -----------------------------------------------------------------------------
Private Sub Sorting()
    'Step 1 Sort
    SortCurves
    
    'Step 2
    Dim q As Long
    Dim MyCurves As New Collection
    Dim MySortingCurve As SortingCurve
    
    For q = 1 To Curves.count
        Set MySortingCurve = New SortingCurve
        With MySortingCurve
            Set .Childs = New Collection
            Set .path = Curves(q)
        End With
        MyCurves.Add MySortingCurve
    Next
    
    'Step 3 IsInner
    SetChilds MyCurves
    
    'Step 4 Set Bloks
    Dim FilmBlocks() As FilmBlock
    Dim UBoundFilmBlock As Long
    Dim c As Long
    UBoundFilmBlock = AllCurvesHeight \ FilmBlokSize
    ReDim FilmBlocks(UBoundFilmBlock)
    For q = 0 To UBoundFilmBlock
        Set FilmBlocks(q).Curves = New Collection
    Next
    For q = 1 To MyCurves.count
        Set MySortingCurve = MyCurves(q)
        c = MySortingCurve.path.Bounds.bottom \ FilmBlokSize
        If c > UBoundFilmBlock Then c = UBoundFilmBlock
        FilmBlocks(c).Curves.Add MySortingCurve
    Next

'    'Step 5
    Dim PreviosPoint As POINTL
    Dim LastAngel As Double
    Dim rc As New Collection
    Dim tempCol As Collection
    
    If lZeroPosition = RightP Then
        PreviosPoint.x = AllCurvesWidth
        LastAngel = 4.712389
    Else
        LastAngel = 1.5707965
    End If
    
    For q = 0 To UBoundFilmBlock
        If FilmBlocks(q).Curves.count > 0 Then
            Set tempCol = FindWay(FilmBlocks(q).Curves, PreviosPoint)
            Set MySortingCurve = tempCol(tempCol.count)
            PreviosPoint = MySortingCurve.path.LastPoint
            AddCollectionToCollection rc, tempCol
        End If
    Next
    Set Curves = SetCurves(rc)
    
    'Step 6
    CornerStep LastAngel
End Sub

' -----------------------------------------------------------------------------
' CornerStep (Sub), source lines 1157-1182
' -----------------------------------------------------------------------------
Private Sub CornerStep(LastAngel As Double)
    Dim q As Long
    Dim path As waPath
    Dim a As Double, aRev As Double
    
    With Curves
        For q = 1 To .count
            Set path = Curves(q)
            With path
                If .closed Then
                    a = AngelStart(path)
                    aRev = AngelEndRev(path)
                    If DeltaAngleAbs(LastAngel, a) > DeltaAngleAbs(LastAngel, aRev) Then .Reverse
                    'If little
                    If .Bounds.bottom - .Bounds.top <= 30000 Then
                        If FindBestAngel(path, LastAngel) Then
                            Curves.Add path, , , q
                            Curves.Remove q
                        End If
                    End If
                End If
            End With
            LastAngel = AngelEnd(path)
        Next
    End With
End Sub

' -----------------------------------------------------------------------------
' FindBestAngel (Function), source lines 1183-1212
' -----------------------------------------------------------------------------
Private Function FindBestAngel(path As waPath, LastAngel As Double) As Boolean
    Dim q As Long
    Dim a As Double, da As Double, BEST_a As Double, BEST_q As Long
    BEST_a = 99999 'Max Val
    Dim segments() As segment
    segments = path.segments

    For q = 1 To UBound(segments)
        With segments(q)
            a = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
            da = DeltaAngleAbs(LastAngel, a)
            If da = 0 Then
                BEST_q = q
                q = UBound(segments)
                GoTo 10
            End If
            If da < BEST_a Then
                BEST_a = da
                BEST_q = q
            End If
10
        End With
    Next
    
    'Change start point
    If BEST_q > 1 Then
        Set path = path.ChangeStartNode(BEST_q)
        FindBestAngel = True
    End If
End Function

' -----------------------------------------------------------------------------
' FindWay (Function), source lines 1213-1265
' -----------------------------------------------------------------------------
Private Function FindWay(SortingCurves As Collection, PreviosPoint As POINTL) As Collection
    Dim NewCurves As New Collection
    Dim sc As SortingCurve
    
    Dim lsc As SortingCurve
    Dim pp As NearestPointInfo
    Dim d1 As Long, d2 As Long
    Dim LastPoint As POINTL
    Dim OldHorisontalDelitel As Long
    
    LastPoint = PreviosPoint
    
    
    With SortingCurves
        Do While .count > 0
            Set sc = NearestSortingCurve(SortingCurves, LastPoint)
            With sc
                If .Childs.count > 0 Then
                    Set .Childs = FindWay(.Childs, LastPoint)
                    
                    Set lsc = .Childs(.Childs.count)
                    LastPoint = lsc.path.LastPoint
                        
                    If .path.closed Then
                        OldHorisontalDelitel = HorisontalDelitel
                        HorisontalDelitel = 1
                        pp = FindNearestPoint(.path, LastPoint.x * HorisontalDelitel, LastPoint.y)
                        HorisontalDelitel = OldHorisontalDelitel
                        Set .path = .path.ChangeStartNode(pp.Number)
                    Else
                        With .path.FirstPoint
                            d1 = Distance(.x, .y, LastPoint.x, LastPoint.y)
                        End With
                        With .path.LastPoint
                            d2 = Distance(.x, .y, LastPoint.x, LastPoint.y)
                        End With
                        If d1 > d2 Then .path.Reverse
                    End If
                End If
                LastPoint = .path.LastPoint
                NewCurves.Add sc
                
            End With
        Loop
    End With
    '====FINAL=========================================================================
    Dim q As Long
    For q = 2 To NewCurves.count
        FinalFinding NewCurves, q, PreviosPoint
    Next
    '==================================================================================
    Set FindWay = NewCurves
End Function

' -----------------------------------------------------------------------------
' FinalFinding (Sub), source lines 1266-1367
' -----------------------------------------------------------------------------
Private Sub FinalFinding(SortingCurves As Collection, CurrentQ As Long, PreviosPoint As POINTL)

    Dim PreviosSC As SortingCurve
    Dim CurrentSC As SortingCurve
    Dim NextSC As SortingCurve
    Dim PathIsClosed As Boolean
    
    Dim Current_FirstPoint As POINTL
    Dim Current_LastPoint As POINTL
    Dim Previos_LastPoint As POINTL
    Dim Next_FirstPoint As POINTL
    
    Set CurrentSC = SortingCurves(CurrentQ)
    Set PreviosSC = SortingCurves(CurrentQ - 1)
    With CurrentSC.path
        PathIsClosed = .closed
        Current_FirstPoint = .FirstPoint
        Current_LastPoint = .LastPoint
    End With
    Previos_LastPoint = PreviosSC.path.LastPoint
    
    Dim q As Long
    Dim a As Long, b_plus_c As Long
    Dim pp As NearestPointInfo
    Dim d1 As Long, d2 As Long
    Dim NeedReverse As Boolean
    Dim SumWay As Long
    
    Dim BEST_SumWay As Long ' BEST_SumWay = 0
    Dim BEST_Before_q As Long
    Dim BEST_Node_q As Long
    Dim BEST_NeedReverse As Boolean
    
    a = Distance(Current_FirstPoint.x, Current_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y)
    
    'First iteration
    Set NextSC = SortingCurves(1)
    Next_FirstPoint = NextSC.path.FirstPoint
    If PathIsClosed Then
        pp = FindNearestPointForTwoShapes(CurrentSC.path, Next_FirstPoint.x, Next_FirstPoint.y, PreviosPoint.x, PreviosPoint.y)
        b_plus_c = pp.Lenth
    Else
        d1 = Distance(Current_FirstPoint.x, Current_FirstPoint.y, PreviosPoint.x, PreviosPoint.y) + Distance(Current_LastPoint.x, Current_LastPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
        d2 = Distance(Current_LastPoint.x, Current_LastPoint.y, PreviosPoint.x, PreviosPoint.y) + Distance(Current_FirstPoint.x, Current_FirstPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
        If d1 > d2 Then
            NeedReverse = True
            b_plus_c = d2
        Else
            NeedReverse = False
            b_plus_c = d1
        End If
    End If
'   Calculate Sum Way
    SumWay = a - b_plus_c
    If SumWay > BEST_SumWay Then
        BEST_SumWay = SumWay
        BEST_Before_q = 1
        BEST_Node_q = pp.Number
        BEST_NeedReverse = NeedReverse
    End If

    For q = 2 To CurrentQ - 1
        Set NextSC = SortingCurves(q)
        Set PreviosSC = SortingCurves(q - 1)
        Next_FirstPoint = NextSC.path.FirstPoint
        Previos_LastPoint = PreviosSC.path.LastPoint
        If PathIsClosed Then
            pp = FindNearestPointForTwoShapes(CurrentSC.path, Next_FirstPoint.x, Next_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y)
            b_plus_c = pp.Lenth
        Else
            d1 = Distance(Current_FirstPoint.x, Current_FirstPoint.y, Previos_LastPoint.x, Previos_LastPoint.y) + Distance(Current_LastPoint.x, Current_LastPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
            d2 = Distance(Current_LastPoint.x, Current_LastPoint.y, Previos_LastPoint.x, Previos_LastPoint.y) + Distance(Current_FirstPoint.x, Current_FirstPoint.y, Next_FirstPoint.x, Next_FirstPoint.y)
            If d1 > d2 Then
                NeedReverse = True
                b_plus_c = d2
            Else
                NeedReverse = False
                b_plus_c = d1
            End If
        End If
'       Calculate Sum Way
        SumWay = a - b_plus_c
        If SumWay > BEST_SumWay Then
            BEST_SumWay = SumWay
            BEST_Before_q = q
            BEST_Node_q = pp.Number
            BEST_NeedReverse = NeedReverse
        End If
    Next

    If BEST_SumWay > 0 Then
        SortingCurves.Remove CurrentQ
        If PathIsClosed Then
            Set CurrentSC.path = CurrentSC.path.ChangeStartNode(BEST_Node_q)
        Else
            If BEST_NeedReverse Then
                CurrentSC.path.Reverse
            End If
        End If
        SortingCurves.Add CurrentSC, , BEST_Before_q
    End If
End Sub

' -----------------------------------------------------------------------------
' FindNearestPointForTwoShapes (Function), source lines 1368-1387
' -----------------------------------------------------------------------------
Private Function FindNearestPointForTwoShapes(path As waPath, x1 As Long, y1 As Long, x2 As Long, y2 As Long) As NearestPointInfo
    Dim q As Long
    Dim d As Long
    Dim segments() As segment
    FindNearestPointForTwoShapes.Lenth = 2147483647
    With path
        segments = .segments
        For q = 1 To UBound(segments)
            With segments(q).P1
                d = Distance(x1, y1, .x, .y) + Distance(x2, y2, .x, .y)
                With FindNearestPointForTwoShapes
                    If d < .Lenth Then
                        .Lenth = d
                        .Number = q
                    End If
                End With
            End With
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' NearestSortingCurve (Function), source lines 1388-1453
' -----------------------------------------------------------------------------
Private Function NearestSortingCurve(CurvesForSearching As Collection, LastPoint As POINTL) As SortingCurve
    Dim CurrentSortingCurve As SortingCurve
    Dim minDistance As Long: minDistance = 2147483647

    Dim minQ As Long
    Dim minNodeNumber As Long
    Dim pp As NearestPointInfo
    Dim d As Long

    Dim q As Long
    Dim lx As Long, ly As Long
    
    With LastPoint
        lx = .x
        ly = .y
    End With
    For q = 1 To CurvesForSearching.count
        Set CurrentSortingCurve = CurvesForSearching(q)
        With CurrentSortingCurve
            If .path.closed Then
                pp = FindNearestPoint(.path, lx, ly)
                With pp
                    If .Lenth < minDistance Then
                        minQ = q
                        minDistance = .Lenth
                        minNodeNumber = .Number
                    End If
                End With
            Else
                With .path.FirstPoint
                    d = Distance(.x, .y, lx, ly)
                End With
                If d < minDistance Then
                    minQ = q
                    minDistance = d
                    minNodeNumber = 0
                End If
                
                With .path.LastPoint
                    d = Distance(.x, .y, lx, ly)
                End With
                If d < minDistance Then
                    minQ = q
                    minDistance = d
                    minNodeNumber = -1
                End If
                
            End If
        End With
    Next
    
    '===
    Set NearestSortingCurve = CurvesForSearching(minQ)
    With NearestSortingCurve
        If .Childs.count = 0 Then
            If minNodeNumber > 0 Then
                Set .path = .path.ChangeStartNode(minNodeNumber)
            ElseIf minNodeNumber < 0 Then
                'Need revers not closed curve
                .path.Reverse
            End If
        End If
    End With
    
    CurvesForSearching.Remove minQ
End Function

' -----------------------------------------------------------------------------
' FindNearestPoint (Function), source lines 1454-1473
' -----------------------------------------------------------------------------
Private Function FindNearestPoint(path As waPath, x As Long, y As Long) As NearestPointInfo
    Dim q As Long
    Dim d As Long
    Dim segments() As segment
    FindNearestPoint.Lenth = 2147483647
    With path
        segments = .segments
        For q = 1 To UBound(segments)
            With segments(q).P1
                d = Distance(x, y, .x, .y)
                With FindNearestPoint
                    If d < .Lenth Then
                        .Lenth = d
                        .Number = q
                    End If
                End With
            End With
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' AngelStart (Function), source lines 1476-1480
' -----------------------------------------------------------------------------
Private Function AngelStart(path As waPath) As Double
    With path.FirstSegment
        AngelStart = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
    End With
End Function

' -----------------------------------------------------------------------------
' AngelEnd (Function), source lines 1481-1489
' -----------------------------------------------------------------------------
Private Function AngelEnd(path As waPath) As Double
    With path.LastSegment
        If .Type = CurveSegment Then
            AngelEnd = Angel(.P3.x, .P3.y, .P4.x, .P4.y)
        Else
            AngelEnd = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' AngelEndRev (Function), source lines 1490-1498
' -----------------------------------------------------------------------------
Private Function AngelEndRev(path As waPath) As Double
    With path.LastSegment
        If .Type = CurveSegment Then
            AngelEndRev = Angel(.P4.x, .P4.y, .P3.x, .P3.y)
        Else
            AngelEndRev = Angel(.P2.x, .P2.y, .P1.x, .P1.y)
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' Angel (Function), source lines 1499-1507
' -----------------------------------------------------------------------------
Private Function Angel(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Double
    If x1 = x2 Then
        If y2 > y1 Then Angel = 3.141593
    Else
        Angel = (y2 - y1) / (x2 - x1)
        Angel = Atn(Angel) + 1.570796
        If (x2 - x1) < 0 Then Angel = Angel + 3.141593
    End If
End Function

' -----------------------------------------------------------------------------
' DeltaAngleAbs (Function), source lines 1508-1516
' -----------------------------------------------------------------------------
Private Function DeltaAngleAbs(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngleAbs = a2 - a1
    If DeltaAngleAbs > 3.141593 Then
        DeltaAngleAbs = DeltaAngleAbs - 6.283186
    ElseIf DeltaAngleAbs < -3.141593 Then
        DeltaAngleAbs = DeltaAngleAbs + 6.283186
    End If
    DeltaAngleAbs = Abs(DeltaAngleAbs)
End Function

' -----------------------------------------------------------------------------
' SetCurves (Function), source lines 1517-1537
' -----------------------------------------------------------------------------
Private Function SetCurves(SortingCurves As Collection) As Collection
    Dim q As Long
    Dim sc As SortingCurve
    Dim c2 As Collection
    Dim NewCurves As New Collection
    
    With SortingCurves
        For q = 1 To .count
            Set sc = SortingCurves(q)
            With sc
                If .Childs.count > 0 Then
                    Set c2 = SetCurves(.Childs)
                    AddCollectionToCollection NewCurves, c2
                End If
                NewCurves.Add .path
            End With
        Next
    End With
    
    Set SetCurves = NewCurves
End Function

' -----------------------------------------------------------------------------
' AddCollectionToCollection (Sub), source lines 1538-1543
' -----------------------------------------------------------------------------
Private Sub AddCollectionToCollection(Col1 As Collection, col2 As Collection)
    Dim q As Long
    For q = 1 To col2.count
        Col1.Add col2(q)
    Next
End Sub

' -----------------------------------------------------------------------------
' Distance (Function), source lines 1544-1549
' -----------------------------------------------------------------------------
Private Function Distance(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Long
    Static a As Double, b As Double
    a = (y2 - y1)
    b = (x2 - x1) \ HorisontalDelitel
    Distance = Sqr(a * a + b * b)
End Function

' -----------------------------------------------------------------------------
' SetChilds (Sub), source lines 1550-1593
' -----------------------------------------------------------------------------
Private Sub SetChilds(CurvesForSetChilds As Collection)

    Dim q As Long, q2 As Long
    Dim MySortingCurve As SortingCurve, MySortingCurve2 As SortingCurve
    Dim left As Long, right As Long, top As Long, bottom As Long
    Dim NeedExit As Boolean

    q = 1
    Do While q < CurvesForSetChilds.count
        Set MySortingCurve = CurvesForSetChilds(q)
        
        With MySortingCurve.path.Bounds
            left = .left
            right = .right
            top = .top
            bottom = .bottom
        End With
        
        NeedExit = False
        
        For q2 = q + 1 To CurvesForSetChilds.count
            Set MySortingCurve2 = CurvesForSetChilds(q2)
            
            With MySortingCurve2.path.Bounds
                If left >= .left Then
                    If right <= .right Then
                        If top >= .top Then
                            If bottom <= .bottom Then
                                MySortingCurve2.Childs.Add MySortingCurve
                                CurvesForSetChilds.Remove q
                                q = q - 1
                                NeedExit = True
                            End If
                        End If
                    End If
                End If
            End With
            
            If NeedExit Then Exit For
        Next
        q = q + 1
    Loop

End Sub

' -----------------------------------------------------------------------------
' SortCurves (Sub), source lines 1595-1628
' -----------------------------------------------------------------------------
Private Sub SortCurves()
    Dim i As Long
    Dim fSwitched As Boolean
    
    Dim count As Long
    Dim path As waPath
    count = Curves.count - 1
    
    Dim sum1 As Long
    Dim sum2 As Long
    
    Do
        fSwitched = False
        For i = 1 To count
            
            Set path = Curves(i)
            With path.Bounds
                sum1 = (.right - .left) + (.bottom - .top)
            End With
            
            Set path = Curves(i + 1)
            With path.Bounds
                sum2 = (.right - .left) + (.bottom - .top)
            End With
            
            If sum1 > sum2 Then
                fSwitched = True
                Set path = Curves(i)
                Curves.Remove i
                Curves.Add path, , , i
            End If
        Next
    Loop While fSwitched
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 1630-1632
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    ReleaseDC lHWndForm, hDC
End Sub
