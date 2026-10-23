Private Sub ClosePathIfDistanceLessThan(SubPath As VGCore.SubPath, ByVal dist As Double)
    With SubPath
        If Not .closed Then
            If SubPath.Nodes.First.GetDistanceFrom(SubPath.Nodes.Last) <= dist Then
                .closed = True
                Exit Sub
            End If
        End If
    End With
End Sub
