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
