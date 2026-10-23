Private Function CheckCOMVidAndPid(VidAndPid As String) As String
    Dim q As Long
    For q = 0 To UBound(Interfaces)
        With Interfaces(q)
            If Not .isComPortAdapter Then
                If UCase(VidAndPid) = UCase(.VidAndPid) Then
                    CheckCOMVidAndPid = VidAndPid
                    Exit Function
                End If
            End If
        End With
    Next
End Function
