Private Function TestShapeForUsedShapes(StaticId As Long, ByRef sr() As ShapeRange, qq As Long) As Boolean
    Dim q As Long, q2 As Long
    For q = 1 To qq
        With sr(q)
            For q2 = 1 To .shapes.count
                If .shapes(q2).StaticId = StaticId Then
                    TestShapeForUsedShapes = True
                    Exit Function
                End If
            Next
        End With
    Next
End Function
