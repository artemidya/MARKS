' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: ConverterCorelCrvToPlotCalcCrv.cls
' Procedures: 9
' =============================================================================

' -----------------------------------------------------------------------------
' GetCurves (Function), source lines 18-60
' -----------------------------------------------------------------------------
Public Function GetCurves(ByRef Height As Long, ByRef Width As Long, ByRef MaxColor As RGBA, Optional ByRef colors As Collection) As Collection

    'Start Process
    Processing = True
    Height = 0
    Width = 0
    Set Curves = New Collection
    Set MyColors = New Collection
    On Error GoTo 20
    If Not Application.ActiveSelectionRange Is Nothing Then
        On Error GoTo 10
        Dim lastUnit As cdrUnit
        lastUnit = Application.ActiveDocument.Unit
        Application.ActiveDocument.Unit = cdrMillimeter
        
        Height = Application.ActiveSelection.SizeHeight * 1000
        Width = Application.ActiveSelection.SizeWidth * 1000

        lastT = Timer
        Percent = 0
        
        Dim mx As New waMatrix
        mx.Translate -Application.ActiveSelection.LeftX * 1000, -Application.ActiveSelection.BottomY * 1000
        GetCurvesCycle Application.ActiveSelectionRange.shapes, mx, True
        
        RaiseEvent OnPercentChange(100)
        
        
        
        
10
        Application.ActiveDocument.Unit = lastUnit
    End If
20
    Set GetCurves = Curves
    Set colors = MyColors
    MaxColor = MaxColorF
     
    Set Curves = Nothing
    Set MyColors = Nothing
    'End Process
    Processing = False
End Function

' -----------------------------------------------------------------------------
' MaxColorF (Function), source lines 61-81
' -----------------------------------------------------------------------------
Private Function MaxColorF() As RGBA

    Dim colorInfo1 As ColorInfo
    Dim MaxColorInfo As New ColorInfo
    
    MaxColorInfo.R = 255
    MaxColorInfo.g = 255
    MaxColorInfo.b = 255
    MaxColorInfo.a = 255
    
    For Each colorInfo1 In MyColors
        If colorInfo1.count > MaxColorInfo.count And colorInfo1.a <> 0 Then
            Set MaxColorInfo = colorInfo1
        End If
    Next
    
    MaxColorF.R = MaxColorInfo.R
    MaxColorF.g = MaxColorInfo.g
    MaxColorF.b = MaxColorInfo.b
    MaxColorF.a = MaxColorInfo.a
End Function

' -----------------------------------------------------------------------------
' GetCurvesCycle (Sub), source lines 82-111
' -----------------------------------------------------------------------------
Private Sub GetCurvesCycle(shapes As VGCore.shapes, mx As waMatrix, Optional UpLevel As Boolean)
    Dim q1 As Long
    For q1 = 1 To shapes.count
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
                GetColor shapes(q1)
                GetCurve .DisplayCurve, mx
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes, mx
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then
                    GetColor shapes(q1)
                    GetCurve .DisplayCurve, mx
                End If
            Else
                'for text
                TryConvertToCurves shapes(q1), mx
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' GetColor (Sub), source lines 112-143
' -----------------------------------------------------------------------------
Private Sub GetColor(shape As VGCore.shape)
    Dim c As VGCore.Color
    With shape.Fill
        Dim Color As RGBA
        If .Type = cdrUniformFill Then
            If .UniformColor.Type = cdrColorRGB Then
                Set c = .UniformColor
            Else
                Set c = .UniformColor.GetCopy
                c.ConvertToRGB
            End If
            Color.R = c.RGBRed
            Color.g = c.RGBGreen
            Color.b = c.RGBBlue
            
            'Transparency
            If shape.Transparency.Type = cdrUniformTransparency Then
                Color.a = (100 - shape.Transparency.Uniform) * 2.55
                If Color.a < 127 Then Color.a = 127
            Else
                Color.a = 255
            End If
            
'        Else
'            Color.R = 255
'            Color.G = 255
'            Color.B = 255
'            Color.A = 255
        End If
        PlusColor Color
    End With
End Sub

' -----------------------------------------------------------------------------
' PlusColor (Sub), source lines 144-163
' -----------------------------------------------------------------------------
Private Sub PlusColor(Color As RGBA)
    Dim q As Long
    Dim colorInfo1 As ColorInfo

    For Each colorInfo1 In MyColors
        If colorInfo1.R = Color.R And colorInfo1.g = Color.g And colorInfo1.b = Color.b And colorInfo1.a = Color.a Then
            colorInfo1.count = colorInfo1.count + 1
            Exit Sub
        End If
    Next

    'Add color
    Set colorInfo1 = New ColorInfo
    colorInfo1.R = Color.R
    colorInfo1.g = Color.g
    colorInfo1.b = Color.b
    colorInfo1.a = Color.a
    colorInfo1.count = 1
    MyColors.Add colorInfo1
End Sub

' -----------------------------------------------------------------------------
' TryConvertToCurves (Sub), source lines 164-179
' -----------------------------------------------------------------------------
Private Sub TryConvertToCurves(shape As VGCore.shape, mx As waMatrix)
    Application.Optimization = True
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        GetColor s2
        GetCurve s2.DisplayCurve, mx
    ElseIf s2.Type = cdrGroupShape Then
        GetCurvesCycle s2.shapes, mx
    End If
10
    s2.Delete
    Application.Optimization = False
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 180-184
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    If shape.DisplayCurve.Nodes.count > 0 Then TryDisplayCurve = True
10
End Function

' -----------------------------------------------------------------------------
' GetCurve (Sub), source lines 185-218
' -----------------------------------------------------------------------------
Private Sub GetCurve(crv As VGCore.curve, mx As waMatrix)
    Dim c As Long:    c = 1000
    With crv
        Dim path As waPath
        Dim q As Long, q2 As Long
        For q = 1 To .Subpaths.count
            If .Subpaths(q).Nodes.count > 1 Then
                With .Subpaths(q)
                    Set path = New waPath
                    path.StartSubPath
                    With .StartNode
                        path.AddLineSegment .PositionX * c, .PositionY * c
                    End With
                    For q2 = 1 To .segments.count
                        With .segments(q2)
                            Select Case .Type
                                Case cdrCurveSegment
                                    path.AddBezierSegment .StartNode.PositionX * c, .StartNode.PositionY * c, .StartingControlPointX * c, .StartingControlPointY * c, .EndingControlPointX * c, .EndingControlPointY * c, .EndNode.PositionX * c, .EndNode.PositionY * c
                                Case cdrLineSegment
                                    path.AddLineSegment .EndNode.PositionX * c, .EndNode.PositionY * c
                                Case cdrMixedSegments
                                    MsgBox "ERR mixed! (cdrMixedSegments)", vbCritical
                                    Exit Sub
                            End Select
                        End With
                    Next
                    If .closed Then path.CloseSubPath
                End With
                path.ApplyMatrixTransform mx
                Curves.Add path
            End If
        Next
    End With
End Sub

' -----------------------------------------------------------------------------
' Break (Sub), source lines 219-221
' -----------------------------------------------------------------------------
Public Sub Break()
    Processing = False
End Sub
