Private Function FindNearestPoint(path As waPath, x As Long, y As Long) As NearestPointInfo
    Dim q As Long
    Dim d As Long
    Dim segments() As segment
    FindNearestPoint.Lenth = 2147483647
    With path
        segments = .segments
        For q = 1 To UBound(segments)
            With segments(q).P1
                d = Distance(x, y, .x, .y)
                With FindNearestPoint
                    If d < .Lenth Then
                        .Lenth = d
                        .Number = q
                    End If
                End With
            End With
        Next
    End With
End Function
