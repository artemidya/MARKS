Private Sub AddToVidAndPid(ByRef lVidAndPid() As String, ByVal NewElement As String)
    If Not Not lVidAndPid() Then
        Dim NewIndex As Long
        NewIndex = UBound(lVidAndPid) + 1
        ReDim Preserve lVidAndPid(NewIndex)
        lVidAndPid(NewIndex) = NewElement
    Else
        ReDim lVidAndPid(0)
        lVidAndPid(0) = NewElement
    End If
End Sub
