Private Function CreateCurveFromShapes(shapes As VGCore.shapes, CurveFromBitmap As curve) As curve
    Set CreateCurveFromShapes = VGCore.ActiveDocument.CreateCurve
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrCurveShape
                Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).curve))
            Case cdrShapeType.cdrGroupShape, cdrShapeType.cdrBlendGroupShape, cdrShapeType.cdrContourGroupShape, cdrArtisticMediaGroupShape
                Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CreateCurveFromShapes(shapes(q).shapes, CurveFromBitmap))
            Case cdrShapeType.cdrBitmapShape
                If shapes(q).Bitmap.Cropped Then
                    Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).Bitmap.CropEnvelope))
                Else
                    With shapes(q)
                        Dim rect2 As curve
                        Set rect2 = VGCore.ActiveDocument.CreateCurve
                        rect2.AppendSubpathRectangle .PositionX, .PositionY, .RightX, .BottomY
                        Set CurveFromBitmap = CurveFromBitmap.WeldWith(rect2)
                    End With
                End If
            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(CloseCurve(shapes(q).DisplayCurve))
                Else
                    With shapes(q)
                        Dim rect As curve
                        Set rect = VGCore.ActiveDocument.CreateCurve
                        rect.AppendSubpathRectangle .PositionX, .PositionY, .RightX, .BottomY
                        Set CreateCurveFromShapes = CreateCurveFromShapes.WeldWith(rect)
                    End With
                End If
         End Select
    Next
End Function
