' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: ConverterCorelCrvToStringL.cls
' Procedures: 7
' =============================================================================

' -----------------------------------------------------------------------------
' GetStringL (Function), source lines 23-45
' -----------------------------------------------------------------------------
Public Function GetStringL(shape As VGCore.shape) As String
    'Start Process
    Processing = True
    If Not shape Is Nothing Then
        On Error GoTo 10
        Dim lastUnit As cdrUnit
        lastUnit = Application.ActiveDocument.Unit
        Application.ActiveDocument.Unit = cdrMillimeter
        lastT = Timer
        Percent = 0

        

        GetCurvesCycle shape.shapes, True
        RaiseEvent OnPercentChange(100)
10
        Application.ActiveDocument.Unit = lastUnit
    End If
    GetStringL = OutString
    OutString = ""
    'End Process
    Processing = False
End Function

' -----------------------------------------------------------------------------
' GetCurvesCycle (Sub), source lines 47-78
' -----------------------------------------------------------------------------
Private Sub GetCurvesCycle(shapes As VGCore.shapes, Optional UpLevel As Boolean)
    Dim q1 As Long
    For q1 = shapes.count To 1 Step -1
        If UpLevel Then
            If Timer - lastT > RefreshTime Then
                Percent = (CSng(q1) / shapes.count) * 100
                lastT = Timer
                RaiseEvent OnPercentChange(Percent)
            End If
        End If
        DoEvents
        If Processing = False Then Exit Sub
        With shapes(q1)
            If .Type = cdrCurveShape Then
                
                
                OutString = OutString & GetColor(shapes(q1))
                GetCurve .DisplayCurve
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then
                    OutString = OutString & GetColor(shapes(q1))
                    GetCurve .DisplayCurve
                End If
            Else
                'for text
                TryConvertToCurves shapes(q1)
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' GetColor (Function), source lines 79-132
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' TryConvertToCurves (Sub), source lines 134-149
' -----------------------------------------------------------------------------
Private Sub TryConvertToCurves(shape As VGCore.shape)
    Application.Optimization = True
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        OutString = OutString & GetColor(s2)
        GetCurve s2.DisplayCurve
    ElseIf s2.Type = cdrGroupShape Then
        GetCurvesCycle s2.shapes
    End If
10
    s2.Delete
    Application.Optimization = False
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 150-154
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    If shape.DisplayCurve.Nodes.count > 0 Then TryDisplayCurve = True
10
End Function

' -----------------------------------------------------------------------------
' GetCurve (Sub), source lines 155-189
' -----------------------------------------------------------------------------
Private Sub GetCurve(crv As VGCore.curve)
    Dim c As Long:    c = 1000
    With crv
'        Dim path As waPath
        Dim q As Long, q2 As Long
        For q = 1 To .Subpaths.count
            With .Subpaths(q)
'                Set path = New waPath
'                path.StartSubPath
                With .StartNode
'                    path.AddLineSegment .PositionX * c, .PositionY * c
                    OutString = OutString & "U" & CLng(.PositionX * c) & "," & CLng(.PositionY * c)
                End With
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        Select Case .Type
                            Case cdrCurveSegment
'                                path.AddBezierSegment .StartNode.PositionX * c, .StartNode.PositionY * c, .StartingControlPointX * c, .StartingControlPointY * c, .EndingControlPointX * c, .EndingControlPointY * c, .EndNode.PositionX * c, .EndNode.PositionY * c
                                 OutString = OutString & "B" & CLng(.StartingControlPointX * c) & "," & CLng(.StartingControlPointY * c) & "," & CLng(.EndingControlPointX * c) & "," & CLng(.EndingControlPointY * c) & "," & CLng(.EndNode.PositionX * c) & "," & CLng(.EndNode.PositionY * c)
                            Case cdrLineSegment
'                                path.AddLineSegment .EndNode.PositionX * c, .EndNode.PositionY * c
                                 OutString = OutString & "D" & CLng(.EndNode.PositionX * c) & "," & CLng(.EndNode.PositionY * c)
                            Case cdrMixedSegments
                                MsgBox "ERR mixed! (cdrMixedSegments)", vbCritical
                                Exit Sub
                        End Select
                    End With
                Next
'                If .Closed Then path.CloseSubPath
            End With
            
'            Curves.Add path
        Next
    End With
End Sub

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 193-195
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Processing = False
End Sub
