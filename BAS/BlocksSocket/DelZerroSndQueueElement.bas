Private Sub DelZerroSndQueueElement()
    If UBound(SndQueueElements()) = 0 Then
        Erase SndQueueElements()
    Else
        Dim q As Long
        For q = 0 To UBound(SndQueueElements()) - 1
            SndQueueElements(q) = SndQueueElements(q + 1)
        Next
        ReDim Preserve SndQueueElements(UBound(SndQueueElements()) - 1)
    End If
End Sub
