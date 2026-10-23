Private Sub AddToSndQueueBytes(ByRef bytes() As Byte, Optional ByVal cp As Long = 0, Optional ByVal needConvertToLng As Long = 0, Optional ByVal Flags As Byte = 0)
    If Not Not SndQueueElements Then
        ReDim Preserve SndQueueElements(UBound(SndQueueElements) + 1)
    Else
        ReDim SndQueueElements(0)
    End If
    With SndQueueElements(UBound(SndQueueElements))
        .ByteBuff() = bytes()
        .cp = cp
        .Flags = Flags
        .needConvertToLng = needConvertToLng
    End With
End Sub
