Private Function DeltaAnglBetweenSegments(Segment1 As VGCore.segment, Segment2 As VGCore.segment) As Double
    On Error GoTo Err
    Static d As Double
    d = Abs(Angel(Segment1) - Angel(Segment2))
    If d > 1.5707963267949 Then
        d = 3.14159265358979 - d
    End If
    DeltaAnglBetweenSegments = d
Err:
End Function
