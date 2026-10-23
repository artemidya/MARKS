' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: ConvertAllToCurves.bas
' Procedures: 8
' =============================================================================

' -----------------------------------------------------------------------------
' Convert (Sub), source lines 4-36
' -----------------------------------------------------------------------------
Public Sub Convert()
    On Error GoTo er:
    UserData.Load
    Dim RegData1 As New RegData
    
    Application.ActiveDocument.PreserveSelection = True
    Application.Optimization = True


    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("ConvertAllToCurves.Command")
    AddLinkedToActiveSelection
    Dim OrigSelection As VGCore.ShapeRange
    Set OrigSelection = ActiveSelectionRangeByOrder()


    Dim convertedShapes As VGCore.ShapeRange
    Set convertedShapes = ConvertShapes(OrigSelection.shapes, RegData1.DelFill, RegData1.DelOutline)

    DeleteUnvisibleLines convertedShapes
er:
    On Error Resume Next
    Application.ActiveDocument.ClearSelection
    OrigSelection.Delete
    If RegData1.UngroupAll Then
        Set convertedShapes = convertedShapes.UngroupAllEx
    End If
    convertedShapes.AddToSelection

    Application.ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.ActiveDocument.PreserveSelection = False
    Application.Refresh
End Sub

' -----------------------------------------------------------------------------
' AddLinkedToActiveSelection (Sub), source lines 37-42
' -----------------------------------------------------------------------------
Private Sub AddLinkedToActiveSelection()
    Dim q As Long
    For q = 1 To VGCore.ActiveSelectionRange.count
        VGCore.ActiveSelectionRange(q).GetLinkedShapes(cdrLinkAllConnections).AddToSelection
    Next
End Sub

' -----------------------------------------------------------------------------
' ActiveSelectionRangeByOrder (Function), source lines 43-54
' -----------------------------------------------------------------------------
Private Function ActiveSelectionRangeByOrder() As VGCore.ShapeRange
    Dim q2 As Long
    Dim newSr As New VGCore.ShapeRange
    With ActiveDocument.ActivePage
        For q2 = 1 To .shapes.count
            If .shapes(q2).Selected Then
                newSr.Add .shapes(q2)
            End If
        Next
    End With
    Set ActiveSelectionRangeByOrder = newSr.ReverseRange
End Function

' -----------------------------------------------------------------------------
' DeleteUnvisibleLines (Sub), source lines 55-66
' -----------------------------------------------------------------------------
Private Sub DeleteUnvisibleLines(sr As VGCore.ShapeRange)
    Dim s As shape
    For Each s In sr
        With s
            If .Type = cdrCurveShape Then
                If Not .curve.closed Then
                    If .Outline.Type = cdrNoOutline Then .Delete
                End If
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' ConvertShapes (Function), source lines 68-140
' -----------------------------------------------------------------------------
Private Function ConvertShapes(shapes As VGCore.shapes, Optional DelFill As Boolean = False, Optional DelOutline As Boolean = False) As VGCore.ShapeRange
    On Error Resume Next
    Dim q As Long
    Dim nShape As VGCore.shape
    Dim sr As VGCore.ShapeRange

    Set ConvertShapes = New VGCore.ShapeRange

    For q = 1 To shapes.count
        Select Case shapes(q).Type
            Case cdrCurveShape
                Set nShape = shapes(q).Layer.CreateCurve(shapes(q).curve)
                CopyPropertiesFrom nShape, shapes(q), DelFill, DelOutline
                ConvertShapes.Add nShape

            Case cdrDropShadowGroupShape
                Set nShape = shapes(q).BreakApartEx.shapes.First
                Set nShape = nShape.ConvertToBitmap
                ConvertShapes.Add nShape

            Case cdrContourGroupShape, cdrArtisticMediaGroupShape, cdrBlendGroupShape, cdrExtrudeGroupShape, cdrCustomEffectGroupShape
                Set sr = shapes(q).BreakApartEx.UngroupEx
                ConvertShapes.AddRange ConvertShapes(sr.shapes, DelFill, DelOutline)
                sr.Delete

            Case cdrGroupShape
                Set sr = shapes(q).Duplicate.UngroupEx
                Set nShape = ConvertShapes(sr.shapes, DelFill, DelOutline).Group
                sr.Delete
                ConvertShapes.Add nShape

            Case cdrShapeType.cdrBitmapShape
                'Do Nothing
                Set nShape = shapes(q).Duplicate
                nShape.OrderToFront
                ConvertShapes.Add nShape

            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set nShape = shapes(q).Layer.CreateCurve(shapes(q).DisplayCurve)
                    CopyPropertiesFrom nShape, shapes(q), DelFill, DelOutline
                    ConvertShapes.Add nShape

                ElseIf TryConvertToCurves(shapes(q)) Then
                    If shapes(q).Type = cdrCurveShape Then
                        Set nShape = shapes(q).Layer.CreateCurve(shapes(q).curve)
                        CopyPropertiesFrom nShape, shapes(q), DelFill, DelOutline
                        ConvertShapes.Add nShape

                    ElseIf shapes(q).Type = cdrGroupShape Then
                        Set sr = shapes(q).Duplicate.UngroupEx
                            If sr.count > 0 Then
                            Set nShape = ConvertShapes(sr.shapes, DelFill, DelOutline).Group
                            sr.Delete
                            ConvertShapes.Add nShape
                        End If
                    Else
                        'Do Nothing
                        Set nShape = shapes(q).Duplicate
                        nShape.OrderToFront
                        ConvertShapes.Add nShape

                    End If
                Else
                    'Do Nothing
                    Set nShape = shapes(q).Duplicate
                    nShape.OrderToFront
                    ConvertShapes.Add nShape

                End If
        End Select
    Next
End Function

' -----------------------------------------------------------------------------
' CopyPropertiesFrom (Sub), source lines 141-166
' -----------------------------------------------------------------------------
Private Sub CopyPropertiesFrom(ShapeTo As VGCore.shape, ShapeFrom As VGCore.shape, DelFill As Boolean, DelOutline As Boolean)
    If DelFill Then
        If ShapeTo.Type = cdrCurveShape Then
            If ShapeTo.curve.closed Then
                If ShapeFrom.Outline.Type <> cdrNoOutline Then
                    If Not DelOutline Then
                        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
                        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
                    End If
                End If
            Else
                If Not DelOutline Or ShapeFrom.Outline.Type = cdrNoOutline Then
                    ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
                    ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
                End If
            End If
        End If
    Else
        ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyFill
        If Not DelOutline Or ShapeFrom.Outline.Type = cdrNoOutline Then
            ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlineColor
            ShapeTo.CopyPropertiesFrom ShapeFrom, cdrCopyOutlinePen
        End If
    End If
    
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 167-174
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
' TryConvertToCurves (Function), source lines 175-180
' -----------------------------------------------------------------------------
Private Function TryConvertToCurves(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    shape.ConvertToCurves
    TryConvertToCurves = True
10
End Function
