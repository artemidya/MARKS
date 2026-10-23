Private Function GetBMPShapeRange(shapes As VGCore.shapes) As VGCore.ShapeRange
    Set GetBMPShapeRange = Application.CreateShapeRange
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrGroupShape
                GetBMPShapeRange.AddRange GetBMPShapeRange(shapes(q).shapes)
            Case cdrShapeType.cdrBitmapShape
                GetBMPShapeRange.Add shapes(q)
         End Select
    Next
End Function
