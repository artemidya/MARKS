' =============================================================================
' Group: 04_Plot_Document_and_Export
' Source module: ConverterPlotCalcCrvToDMPL.cls
' Procedures: 16
' =============================================================================

' -----------------------------------------------------------------------------
' ConvertPreRun (Function), source lines 24-110
' -----------------------------------------------------------------------------
Public Function ConvertPreRun(Profile As Profile, ByVal dx As Long, ByVal dy As Long, ByVal AllCurvesHeight As Long, ByVal AllCurvesWidth As Long) As String
    Last_X = -2147483647
    Last_Y = -2147483647
    
    Dim x() As Long, y() As Long
    Dim kx As Double, ky As Double
    Dim xp As Long, yp As Long

    ConvertPreRun = ";:HL0,EC1,P1,V5,A"
       
    With Profile
        Select Case .ZeroPosition
            Case ZeroPositions.RightP
                kx = 0.04 * Abs(.HeightScale)
                ky = 0.04 * Abs(.WidthScale)
                
                ReDim x(4), y(4)
                x(0) = 0
                y(0) = 0
                
                x(1) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                y(1) = 0
                
                x(2) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                y(2) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                
                x(3) = 0
                y(3) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                
                x(4) = 0
                y(4) = 0
            Case ZeroPositions.CenterP
                kx = 0.04 * Abs(.WidthScale)
                ky = 0.04 * Abs(.HeightScale)
                
                ReDim x(5), y(5)
                x(0) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(0) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(1) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(1) = (AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(2) = -(AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(2) = (AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(3) = -(AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(3) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(4) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(4) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(5) = 0
                y(5) = 0
            Case ZeroPositions.LeftP
                kx = 0.04 * Abs(.WidthScale)
                ky = 0.04 * Abs(.HeightScale)
                
                ReDim x(4), y(4)
                x(0) = 0
                y(0) = 0
                
                x(1) = 0
                y(1) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                
                x(2) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                y(2) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                
                x(3) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                y(3) = 0
                
                x(4) = 0
                y(4) = 0
        End Select
    End With
    
    Dim q As Long
    For q = 0 To UBound(x) - 1
        xp = x(q) * kx + dx
        yp = y(q) * ky + dy
        ConvertPreRun = ConvertPreRun & "U" & xp & "," & yp & ","
        If Profile.CornerPoints Then ConvertPreRun = ConvertPreRun & "D"
    Next
    xp = x(q) * kx + dx
    yp = y(q) * ky + dy
    ConvertPreRun = ConvertPreRun & "U" & xp & "," & yp & ","
    ConvertPreRun = ConvertPreRun & "U@"
End Function

' -----------------------------------------------------------------------------
' Convert (Function), source lines 111-182
' -----------------------------------------------------------------------------
Public Function Convert(Curves As Collection, Profile As Profile, ByVal dx As Long, ByVal dy As Long, ByVal AllCurvesHeight As Long, ByVal AllCurvesWidth As Long) As String
    Last_X = -2147483647
    Last_Y = -2147483647
    LastCommandIsPD = False
    
    Set MyCurves = CloneCollection(Curves)
    Dim path As waPath
    Dim path2 As waPath
    Dim mx As New waMatrix
            
    'LastAngel =
    Select Case Profile.ZeroPosition
        Case ZeroPositions.RightP
            mx.Rotate 270
            mx.Translate 0, AllCurvesWidth + 1
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment AllCurvesWidth, 0
                path.AddLineSegment AllCurvesWidth - Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment AllCurvesWidth, AllCurvesHeight + 1000
                path2.AddLineSegment AllCurvesWidth - Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 3.14159265358979
        Case ZeroPositions.CenterP
            mx.Translate -AllCurvesWidth / 2, -AllCurvesHeight / 2
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment 0, 0
                path.AddLineSegment Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment 0, AllCurvesHeight + 1000
                path2.AddLineSegment Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 1.5707963267949
        Case ZeroPositions.LeftP
            If Profile.KnifeOffset > 0 Then
                'Add KnifeOffset starting line if need
                Set path = New waPath
                path.StartSubPath
                path.AddLineSegment 0, 0
                path.AddLineSegment Profile.KnifeOffset * 3, 0
                MyCurves.Add path, , 1
                
                Set path2 = New waPath
                path2.StartSubPath
                path2.AddLineSegment 0, AllCurvesHeight + 1000
                path2.AddLineSegment Profile.KnifeOffset * 3, AllCurvesHeight + 1000
                MyCurves.Add path2
            End If
            LastAngel = 1.5707963267949
    End Select
    
    'Apply Matrix
    For Each path In MyCurves
        path.ApplyMatrixTransform mx
    Next

    Convert = CurvesToString(dx, dy, Profile)
End Function

' -----------------------------------------------------------------------------
' CurvesToString (Function), source lines 183-280
' -----------------------------------------------------------------------------
Private Function CurvesToString(ByVal dx As Long, ByVal dy As Long, Profile As Profile) As String
    Dim q As Long
    Dim path As waPath
    Dim segments() As segment
    
    Dim flatness As Single
    Dim kx As Double, ky As Double

    Dim RelativeCoordinates As Boolean
    
    Dim VS1 As Long, VS2 As Long
    
    If Profile.CoordinateSystem = Relative Then RelativeCoordinates = True

    With Profile
        If .ZeroPosition = RightP Then
            kx = 0.04 * Abs(.HeightScale)
            ky = 0.04 * Abs(.WidthScale)
        Else
            kx = 0.04 * Abs(.WidthScale)
            ky = 0.04 * Abs(.HeightScale)
        End If

        If .ZeroPosition <> CenterP Then
            dx = dx + .KnifeOffset * kx
            dy = dy + .KnifeOffset * ky
        End If
        
        CurvesToString = CurvesToString & .Initializer
        CurvesToString = CurvesToString & .BeginWorkFile
        
'        If .KnifeOffset > 0 Then
'            CurvesToString = CurvesToString & "TO0;"
'        End If
        
'        If .ActiveAcceleration > 0 Then
'            CurvesToString = CurvesToString & "AS" & Replace(Format(.Accelerations(.ActiveAcceleration - 1) / 10 * 2 / 9.80665, "###0.0"), ",", ".") & ";"
'        End If
'
        If .ActiveForce > 0 Then
            CurvesToString = CurvesToString & "BP" & .ActiveForce & ","
        End If
              
        If .ActiveSpeed > 0 Then
            VS1 = .Speeds(.ActiveSpeedUP - 1)
            VS2 = .Speeds(.ActiveSpeed - 1)
            If Not .SeparateParametrs Then VS1 = VS2
        End If
        
        If Not RelativeCoordinates Then CurvesToString = CurvesToString & "A"
        

        If .ApproximationLevel = High Then
            flatness = 2
        ElseIf .ApproximationLevel = Normal Then
            flatness = 6
        Else
            flatness = 12
        End If

        For q = 1 To MyCurves.count
            Set path = MyCurves(q)
            segments = path.segments
            CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, 0, VS1, VS2)
                       
            If .DoublePass Then
                If .KnifeOffset > 0 Then
                    If q > 1 And q < MyCurves.count Then
                        If path.closed Then
                            CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, 2147483647, VS1, VS2)
                        Else
                            CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, 0, VS1, VS2)
                        End If
                    End If
                Else
                    If path.closed Then
                        CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, 2147483647, VS1, VS2)
                    Else
                        CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, 0, VS1, VS2)
                    End If
                End If
            Else
                If .ClosingCurve > 0 Then
                    'If (path.Bounds.bottom - path.Bounds.top) > 30001 Then
                        If path.closed Then
                            'To Closing Curve For figures > 30 mm==============================
                            CurvesToString = CurvesToString & CreateHPGLPath(segments, .KnifeOffset, flatness, dx, dy, kx, ky, path, RelativeCoordinates, .ClosingCurve, VS1, VS2)
                        End If
                    'End If
                End If
            End If
'            If .OneCommandOnPath Then CurvesToString = CurvesToString
        Next
    
        CurvesToString = CurvesToString & .EndWorkFile
        CurvesToString = CurvesToString & .Terminator
    End With
End Function

' -----------------------------------------------------------------------------
' SetVS (Function), source lines 281-287
' -----------------------------------------------------------------------------
Private Function SetVS(vs As Long) As String
    Static lastVS As Long
    If vs <> lastVS Then
        SetVS = "V" & Int(vs / 10) & ","
        lastVS = vs
    End If
End Function

' -----------------------------------------------------------------------------
' CreateHPGLPath (Function), source lines 288-399
' -----------------------------------------------------------------------------
Private Function CreateHPGLPath(segments() As segment, KnifeOffset As Long, flatness As Single, dx As Long, dy As Long, kx As Double, ky As Double, path As waPath, ByRef RelativeCoordinates As Boolean, prolongLength As Long, VS1 As Long, VS2 As Long) As String
    Dim q2 As Long, q3 As Long
    Dim tmpStr As String
    Dim SegmentsFromPolyline() As segment
    Dim ProlongSegment As segment
    Dim x As Long, y As Long
    Dim a As Double, delta As Double, newAngel As Double
    Dim p As POINTL
    Dim sumLegth As Long
          
    If prolongLength = 0 Then
        With segments(0)
            p = PointsFromAngel(LastAngel, KnifeOffset, .P2.x, .P2.y)
        End With
        x = p.x * kx + dx
        y = p.y * ky + dy
        tmpStr = tmpStr & SetVS(VS1)
        If RelativeCoordinates Then
            tmpStr = tmpStr & "A" & PU(x, y, False) & "R"
        Else
            tmpStr = tmpStr & PU(x, y, False)
        End If
        tmpStr = tmpStr & SetVS(VS2)
    End If
    
    For q2 = 1 To UBound(segments)
        With segments(q2)
            Select Case .Type
                Case SegmentType.LineSegment
                    a = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
                    delta = DeltaAngle(LastAngel, a)
                    If Abs(delta) > 1.58 Then
                        newAngel = LastAngel + delta / 2
                        p = PointsFromAngel(newAngel, KnifeOffset, .P1.x, .P1.y)
                        x = p.x * kx + dx
                        y = p.y * ky + dy
                        tmpStr = tmpStr & PD(x, y, RelativeCoordinates)
                    End If
                    
                    'Prolong Line
                    ProlongSegment = ShiftSegment(segments(q2), KnifeOffset)
                    
                    If prolongLength > 0 Then
                        With segments(q2)
                            sumLegth = sumLegth + (Distance(.P1, .P2))
                            If sumLegth > prolongLength Then
                                ProlongSegment = ShiftSegmentNodeP2_Only(ProlongSegment, prolongLength - sumLegth)
                                q2 = UBound(segments) 'for Exit for
                            End If
                        End With
                    End If
                    
                     
                    With ProlongSegment
                        x = .P1.x * kx + dx
                        y = .P1.y * ky + dy
                        tmpStr = tmpStr & PD(x, y, RelativeCoordinates)
                        x = .P2.x * kx + dx
                        y = .P2.y * ky + dy
                        tmpStr = tmpStr & PD(x, y, RelativeCoordinates)
                    End With
                    LastAngel = a
                    
                Case SegmentType.CurveSegment
                    SegmentsFromPolyline = path.GetPolyline(segments(q2), flatness)
                    
                    With SegmentsFromPolyline(1)
                        a = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
                        delta = DeltaAngle(LastAngel, a)
                        If Abs(delta) > 1.58 Then
                            newAngel = LastAngel + delta / 2
                            p = PointsFromAngel(newAngel, KnifeOffset, .P1.x, .P1.y)
                            x = p.x * kx + dx
                            y = p.y * ky + dy
                            tmpStr = tmpStr & PD(x, y, RelativeCoordinates)
                        End If
                    End With
                    
                    SegmentsFromPolyline = ShiftSegments(SegmentsFromPolyline(), KnifeOffset)
                    
                    'Obrezat
                    If prolongLength > 0 Then
                        For q3 = 1 To UBound(SegmentsFromPolyline)
                            With SegmentsFromPolyline(q3)
                                sumLegth = sumLegth + (Distance(.P1, .P2))
                            End With
                            If sumLegth > prolongLength Then
                                SegmentsFromPolyline(q3) = ShiftSegmentNodeP2_Only(SegmentsFromPolyline(q3), prolongLength - sumLegth)
                                ReDim Preserve SegmentsFromPolyline(q3)
                                q2 = UBound(segments) 'for Exit for
                                Exit For
                            End If
                        Next
                    End If
                    
                    'Arcs
                    
                    For q3 = 1 To UBound(SegmentsFromPolyline)
                        With SegmentsFromPolyline(q3)
                            x = .P2.x * kx + dx
                            y = .P2.y * ky + dy
                            tmpStr = tmpStr & PD(x, y, RelativeCoordinates)
                        End With
                    Next
                    With SegmentsFromPolyline(q3 - 1)
                        LastAngel = Angel(.P1.x, .P1.y, .P2.x, .P2.y)
                    End With
            End Select
        End With
    Next
    CreateHPGLPath = tmpStr
End Function

' -----------------------------------------------------------------------------
' Distance (Function), source lines 400-418
' -----------------------------------------------------------------------------
Private Function Distance(P1 As POINTL, P2 As POINTL) As Long
    
    Static x1 As Long, y1 As Long
    Static x2 As Long, y2 As Long
    
    With P1
        x1 = .x
        y1 = .y
    End With
    With P2
        x2 = .x
        y2 = .y
    End With
    
    Static a As Double, b As Double
    a = (y2 - y1)
    b = (x2 - x1)
    Distance = Sqr(a * a + b * b)
End Function

' -----------------------------------------------------------------------------
' PD (Function), source lines 419-428
' -----------------------------------------------------------------------------
Private Function PD(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If Last_X <> x Or Last_Y <> y Then
        If LastCommandIsPD Then
            PD = Pair(x, y, RelativeCoordinates)
        Else
            PD = "D" & Pair(x, y, RelativeCoordinates)
        End If
        LastCommandIsPD = True
    End If
End Function

' -----------------------------------------------------------------------------
' PU (Function), source lines 429-434
' -----------------------------------------------------------------------------
Private Function PU(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If Last_X <> x Or Last_Y <> y Then
        PU = "U" & Pair(x, y, RelativeCoordinates)
    End If
    LastCommandIsPD = False
End Function

' -----------------------------------------------------------------------------
' Pair (Function), source lines 435-443
' -----------------------------------------------------------------------------
Private Function Pair(ByRef x As Long, ByRef y As Long, ByRef RelativeCoordinates As Boolean) As String
    If RelativeCoordinates Then
        Pair = x - Last_X & "," & y - Last_Y & ","
    Else
        Pair = x & "," & y & ","
    End If
    Last_X = x
    Last_Y = y
End Function

' -----------------------------------------------------------------------------
' CloneCollection (Function), source lines 444-454
' -----------------------------------------------------------------------------
Private Function CloneCollection(col As Collection) As Collection
    Set CloneCollection = New Collection
    Dim path As waPath
    Dim newPath As waPath
    Dim q As Long
    For q = 1 To col.count
        Set path = col(q)
        Set newPath = path.Clone
        CloneCollection.Add newPath
    Next
End Function

' -----------------------------------------------------------------------------
' ShiftSegments (Function), source lines 455-461
' -----------------------------------------------------------------------------
Private Function ShiftSegments(segments() As segment, ByVal Shift As Long) As segment()
    Dim q As Long
    For q = 1 To UBound(segments)
        segments(q) = ShiftSegment(segments(q), Shift)
    Next
    ShiftSegments = segments
End Function

' -----------------------------------------------------------------------------
' ShiftSegment (Function), source lines 462-489
' -----------------------------------------------------------------------------
Private Function ShiftSegment(segment As segment, ByVal Shift As Long) As segment
    Dim k As Double
    Dim dx As Long, dy As Long
    Dim sx As Long, sy As Long
    With segment
        dx = (.P2.x - .P1.x)
        dy = (.P2.y - .P1.y)
        If dx = 0 Then
            If dy < 0 Then sy = -Shift Else sy = Shift
        Else
            k = dy / dx
            If dx < 0 Then Shift = -Shift
            sx = Shift / Sqr(1 + k * k)
            sy = k * sx
        End If
    End With
    
    With ShiftSegment
        With .P1
            .x = segment.P1.x + sx
            .y = segment.P1.y + sy
        End With
        With .P2
            .x = segment.P2.x + sx
            .y = segment.P2.y + sy
        End With
    End With
End Function

' -----------------------------------------------------------------------------
' ShiftSegmentNodeP2_Only (Function), source lines 490-513
' -----------------------------------------------------------------------------
Private Function ShiftSegmentNodeP2_Only(segment As segment, ByVal Shift As Long) As segment
    Dim k As Double
    Dim dx As Long, dy As Long
    Dim sx As Long, sy As Long
    With segment
        dx = (.P2.x - .P1.x)
        dy = (.P2.y - .P1.y)
        If dx = 0 Then
            If dy < 0 Then sy = -Shift Else sy = Shift
        Else
            k = dy / dx
            If dx < 0 Then Shift = -Shift
            sx = Shift / Sqr(1 + k * k)
            sy = k * sx
        End If
    End With
    
    ShiftSegmentNodeP2_Only.P1 = segment.P1
    With ShiftSegmentNodeP2_Only.P2
        .x = segment.P2.x + sx
        .y = segment.P2.y + sy
    End With

End Function

' -----------------------------------------------------------------------------
' PointsFromAngel (Function), source lines 514-520
' -----------------------------------------------------------------------------
Private Function PointsFromAngel(ByVal Angel As Double, R As Long, Optional dx As Long, Optional dy As Long) As POINTL
    Angel = Angel - 1.5707963267949
    With PointsFromAngel
        .y = R * Sin(Angel) + dy
        .x = R * Cos(Angel) + dx
    End With
End Function

' -----------------------------------------------------------------------------
' Angel (Function), source lines 521-529
' -----------------------------------------------------------------------------
Private Function Angel(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Double
    If x1 = x2 Then
        If y2 > y1 Then Angel = 3.14159265358979
    Else
        Angel = (y2 - y1) / (x2 - x1)
        Angel = Atn(Angel) + 1.5707963267949
        If (x2 - x1) < 0 Then Angel = Angel + 3.14159265358979
    End If
End Function

' -----------------------------------------------------------------------------
' DeltaAngle (Function), source lines 530-537
' -----------------------------------------------------------------------------
Private Function DeltaAngle(ByRef a1 As Double, ByRef a2 As Double) As Double
    DeltaAngle = a2 - a1
    If DeltaAngle > 3.14159265358979 Then
        DeltaAngle = DeltaAngle - 6.28318530717959
    ElseIf DeltaAngle < -3.14159265358979 Then
        DeltaAngle = DeltaAngle + 6.28318530717959
    End If
End Function
