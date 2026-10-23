Private Function OneOfSubpathsIsOpened(curve As VGCore.curve) As Boolean
    Dim q As Long
    With curve
        For q = 1 To .Subpaths.count
            If Not .Subpaths(q).closed Then
                OneOfSubpathsIsOpened = True
                Exit For
            End If
        Next
    End With
End Function
