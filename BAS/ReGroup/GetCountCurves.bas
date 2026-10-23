Private Sub GetCountCurves(shapes As VGCore.shapes)
    Dim q As Long
    For q = 1 To shapes.count
        With shapes(q)
            If .Type = cdrCurveShape Then
               CountShapes = CountShapes + 1
            ElseIf .Type = cdrGroupShape Then
               GetCountCurves .shapes
            End If
        End With
    Next
End Sub
