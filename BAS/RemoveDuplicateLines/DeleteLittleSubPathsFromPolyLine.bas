Private Sub DeleteLittleSubPathsFromPolyLine(PolyLineSubPaths As VGCore.Subpaths, minLength As Double)
    Dim q As Long
    For q = PolyLineSubPaths.count To 1 Step -1
        With PolyLineSubPaths(q)
            If .Length < minLength Then .Delete
        End With
    Next
End Sub
