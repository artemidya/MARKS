Private Function GetShapesRage(ShapeRage As VGCore.ShapeRange) As VGCore.ShapeRange
    Dim GSR As New VGCore.ShapeRange
    Dim q As Long
    For q = 1 To ShapeRage.count
        With ShapeRage
            Select Case .shapes(q).Type
                Case cdrRectangleShape, cdrEllipseShape, cdrPolygonShape, cdrPerfectShape
                    .shapes(q).ConvertToCurves
                    GSR.Add .shapes(q)
                Case cdrShapeType.cdrCurveShape
                    GSR.Add .shapes(q)
                Case cdrShapeType.cdrGroupShape
                    GSR.AddRange GetShapesRage(.shapes(q).shapes.All)
            End Select
        End With
    Next
    Set GetShapesRage = GSR
End Function
