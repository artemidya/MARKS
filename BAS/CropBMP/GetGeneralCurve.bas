Private Function GetGeneralCurve(shapes As VGCore.shapes) As VGCore.curve
    Set GetGeneralCurve = Application.CreateCurve
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrCurveShape
                Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(shapes(q).curve))
            Case cdrShapeType.cdrGroupShape, cdrShapeType.cdrBlendGroupShape, cdrShapeType.cdrContourGroupShape, cdrArtisticMediaGroupShape
                Set GetGeneralCurve = GetGeneralCurve.WeldWith(GetGeneralCurve(shapes(q).shapes))
            Case cdrShapeType.cdrBitmapShape
                'Do Nothing
            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(shapes(q).DisplayCurve))
                Else
                    Dim NewCrv As VGCore.curve
                    If TryConvertToCurves(shapes(q), NewCrv) Then
                        Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(NewCrv))
                    End If
                End If
         End Select
    Next
End Function
