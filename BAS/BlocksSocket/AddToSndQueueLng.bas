Private Sub AddToSndQueueLng(ByRef lngBuffer() As Long)
    If Not Not SndQueueElements Then
        ReDim Preserve SndQueueElements(UBound(SndQueueElements) + 1)
    Else
        ReDim SndQueueElements(0)
    End If
    SndQueueElements(UBound(SndQueueElements)).LngBuff = lngBuffer
End Sub
