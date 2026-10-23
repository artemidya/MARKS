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
