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
