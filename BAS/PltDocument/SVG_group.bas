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
