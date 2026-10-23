Private Function FindNearestPointForTwoShapes(path As waPath, x1 As Long, y1 As Long, x2 As Long, y2 As Long) As NearestPointInfo
    Dim q As Long
    Dim d As Long
    Dim segments() As segment
    FindNearestPointForTwoShapes.Lenth = 2147483647
    With path
        segments = .segments
        For q = 1 To UBound(segments)
            With segments(q).P1
                d = Distance(x1, y1, .x, .y) + Distance(x2, y2, .x, .y)
                With FindNearestPointForTwoShapes
                    If d < .Lenth Then
                        .Lenth = d
                        .Number = q
                    End If
                End With
            End With
        Next
    End With
End Function
