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
