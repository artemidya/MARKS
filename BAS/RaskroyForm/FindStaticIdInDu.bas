Private Function FindStaticIdInDu(StaticId As Long) As Boolean
    Dim q As Long
    Dim q2 As Long
    For q = 1 To UBound(Du)
        With Du(q)
            For q2 = 1 To .shapes.count
                If .shapes(q2).StaticId = StaticId Then
                    FindStaticIdInDu = True
                    Exit Function
                End If
            Next
        End With
    Next
End Function
