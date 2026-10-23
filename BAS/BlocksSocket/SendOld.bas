Public Sub SendOld(ByRef bytes() As Byte)
    If Not Not bytes Then
        If Not Not SndQueueElements Then
            AddToSndQueueBytes bytes()
        Else
            If SendLng(lSocket, bytes(0), UBound(bytes) + 1, 0) < 1 Then
                AddToSndQueueBytes bytes()
            Else
                'Add Statistic
                AddToStatisticSent UBound(bytes) + 1
            End If
        End If
    End If
End Sub
