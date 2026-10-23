Private Function OneOfSubpathsIsClosed(shape As VGCore.shape) As Boolean
    With shape
        If .Type = cdrCurveShape Then
            Dim q As Long
            With .curve
                For q = 1 To .Subpaths.count
                    If .Subpaths(q).closed Then
                        OneOfSubpathsIsClosed = True
                        Exit For
                    End If
                Next
            End With
        Else
            OneOfSubpathsIsClosed = True
        End If
    End With
End Function
