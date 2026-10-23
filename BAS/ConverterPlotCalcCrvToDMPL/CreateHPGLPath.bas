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
