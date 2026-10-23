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
