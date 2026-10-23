Private Sub DeleteZerro(Curves() As VGCore.curve)
    Dim q As Long, q2 As Long
    For q = 0 To UBound(Curves)
        With Curves(q)
            If .segments.count > 0 Then
                For q2 = .Subpaths.count To 1 Step -1
                    With .Subpaths(q2)
                        If .Length = 0 Then .Delete
                    End With
                Next
            End If
        End With
    Next
End Sub
