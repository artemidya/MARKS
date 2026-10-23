Private Function BoundingBoxCrossTest(ByRef bBox As VGCore.rect, ByRef left As Double, ByRef right As Double, ByRef top As Double, ByRef bottom As Double) As Boolean
    With bBox
        If left <= .right Then
            If right >= .left Then
                If top >= .bottom Then
                    If bottom <= .top Then
                        BoundingBoxCrossTest = True
                    End If
                End If
            End If
        End If
    End With
End Function
