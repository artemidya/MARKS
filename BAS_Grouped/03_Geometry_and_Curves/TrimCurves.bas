' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: TrimCurves.bas
' Procedures: 8
' =============================================================================

' -----------------------------------------------------------------------------
' TrimCrv (Sub), source lines 7-64
' -----------------------------------------------------------------------------
Public Sub TrimCrv()
    On Error GoTo er:
    UserData.Load
    Application.ActiveDocument.PreserveSelection = True
    Application.Optimization = True

    Dim q As Long
    Set GdelShapeRange = New VGCore.ShapeRange
    Set GorderBackShapeRange = New VGCore.ShapeRange
    
    Dim lastUnit As cdrUnit
    With Application.ActiveDocument
        lastUnit = .Unit
        .Unit = cdrMillimeter
    End With
    
    
    
    
    Dim RegData1 As New RegData
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("TrimCurves.Command") & " (" & CDbl(RegData1.TrimOffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    BreakDistance = RegData1.TrimOffset / 1000
    If BreakDistance = 0 Then BreakDistance = -1
    
    Dim OrigSelection As VGCore.ShapeRange
    Set OrigSelection = VGCore.ActiveSelectionRange
    TrimShapes OrigSelection.shapes
    
    
    
    On Error Resume Next
    For q = 1 To GorderBackShapeRange.count
        If GdelShapeRange(q).ParentGroup Is Nothing Then
            GorderBackShapeRange(q).OrderBackOf GdelShapeRange(q)
        ElseIf GdelShapeRange(q).ParentGroup.Type = cdrGroupShape Then
            GorderBackShapeRange(q).OrderBackOf GdelShapeRange(q)
        End If
    Next
    
    
er:
    'Kill shapes
    GdelShapeRange.All.Delete
    
    Application.ActiveDocument.ClearSelection
    OrigSelection.All.AddToSelection
    For q = 1 To GorderBackShapeRange.count
        If GorderBackShapeRange(q).ParentGroup Is Nothing Then
            GorderBackShapeRange(q).AddToSelection
        End If
    Next
    Application.ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.ActiveDocument.PreserveSelection = False
    Application.Refresh
    
    Application.ActiveDocument.Unit = lastUnit
End Sub

' -----------------------------------------------------------------------------
' TrimShapes (Sub), source lines 66-93
' -----------------------------------------------------------------------------
Private Sub TrimShapes(shapes As VGCore.shapes)
    On Error Resume Next
    Dim q As Long
    Dim nShape As VGCore.shape
    
    For q = 1 To shapes.count
        Select Case shapes(q).Type
            Case cdrShapeType.cdrCurveShape
                Set nShape = CreateShape(shapes(q).curve)
                GorderBackShapeRange.Add nShape
                GdelShapeRange.Add shapes(q)
            Case cdrShapeType.cdrGroupShape, cdrShapeType.cdrBlendGroupShape, cdrShapeType.cdrContourGroupShape, cdrArtisticMediaGroupShape
                TrimShapes shapes(q).shapes
            Case cdrShapeType.cdrBitmapShape
                'Do Nothing
            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set nShape = CreateShape(shapes(q).DisplayCurve)
                    GorderBackShapeRange.Add nShape
                    GdelShapeRange.Add shapes(q)
                ElseIf TryConvertToCurves(shapes(q)) Then
                    Set nShape = CreateShape(shapes(q).curve)
                    GorderBackShapeRange.Add nShape
                    GdelShapeRange.Add shapes(q)
                End If
        End Select
    Next
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 94-101
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As VGCore.shape) As Boolean
    On Error GoTo Err
    Dim c As Long
    c = shape.DisplayCurve.Subpaths.count
    TryDisplayCurve = True
    Exit Function
Err:
End Function

' -----------------------------------------------------------------------------
' TryConvertToCurves (Function), source lines 102-111
' -----------------------------------------------------------------------------
Private Function TryConvertToCurves(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    shape.ConvertToCurves
    If shape.Type = cdrCurveShape Then
        TryConvertToCurves = True
    ElseIf shape.Type = cdrGroupShape Then
        TrimShapes shape.shapes
    End If
10
End Function

' -----------------------------------------------------------------------------
' CreateShape (Function), source lines 113-125
' -----------------------------------------------------------------------------
Private Function CreateShape(crv As VGCore.curve) As VGCore.shape
    Set CreateShape = VGCore.ActiveLayer.CreateCurve(crv)
    Dim q As Long
    With CreateShape.curve
        For q = 1 To .Subpaths.count
            ClosePathIfDistanceLessThan .Subpaths(q), BreakDistance * 0.9
            If .Subpaths(q).closed = True And .Subpaths(q).Length > BreakDistance * 1.1 Then
                MaxNodeInSubPath(.Subpaths(q)).BreakApart
                TrimSubPath .Subpaths(q), BreakDistance / 2
            End If
        Next
    End With
End Function

' -----------------------------------------------------------------------------
' TrimSubPath (Sub), source lines 127-152
' -----------------------------------------------------------------------------
Private Sub TrimSubPath(SubPath As VGCore.SubPath, ByVal dist As Double)
    On Error GoTo ex
    With SubPath
        Dim sn As VGCore.Node
        Dim EN As VGCore.Node
        Dim q As Long
        
        If .segments.First.Type = cdrLineSegment Then
            Set sn = .AddNodeAt(dist * 2, cdrAbsoluteSegmentOffset)
            For q = 1 To sn.index
                .Nodes.First.Delete
            Next
        Else
            Set sn = .AddNodeAt(dist, cdrAbsoluteSegmentOffset)
            Set EN = .AddNodeAt(.Length - dist, cdrAbsoluteSegmentOffset)
            For q = 1 To .Nodes.count - EN.index - 1
                .Nodes.Last.Delete
            Next
    
            For q = 1 To sn.index
                .Nodes.First.Delete
            Next
        End If
    End With
ex:
End Sub

' -----------------------------------------------------------------------------
' ClosePathIfDistanceLessThan (Sub), source lines 153-162
' -----------------------------------------------------------------------------
Private Sub ClosePathIfDistanceLessThan(SubPath As VGCore.SubPath, ByVal dist As Double)
    With SubPath
        If Not .closed Then
            If SubPath.Nodes.First.GetDistanceFrom(SubPath.Nodes.Last) <= dist Then
                .closed = True
                Exit Sub
            End If
        End If
    End With
End Sub

' -----------------------------------------------------------------------------
' MaxNodeInSubPath (Function), source lines 163-193
' -----------------------------------------------------------------------------
Private Function MaxNodeInSubPath(SubPath As VGCore.SubPath) As VGCore.Node
    With SubPath
        If .Nodes.count > 1 Then
            Dim maxY As Double
            Dim q As Long
            Set MaxNodeInSubPath = .Nodes(1)
            With .segments(1)
                If .Type = cdrLineSegment Then
                    maxY = (.StartNode.PositionY + .EndNode.PositionY) / 2
                Else
                    maxY = .StartNode.PositionY
                End If
            End With
            For q = 2 To .segments.count
                With .segments(q)
                    If .Type = cdrLineSegment Then
                        If (.StartNode.PositionY + .EndNode.PositionY) / 2 > maxY Then
                            maxY = (.StartNode.PositionY + .EndNode.PositionY) / 2
                            Set MaxNodeInSubPath = .StartNode
                        End If
                    Else
                        If .StartNode.PositionY > maxY Then
                            maxY = .StartNode.PositionY
                            Set MaxNodeInSubPath = .StartNode
                        End If
                    End If
                End With
            Next
        End If
    End With
End Function
