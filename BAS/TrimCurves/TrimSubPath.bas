Private Sub TrimSubPath(SubPath As VGCore.SubPath, ByVal dist As Double)
    On Error GoTo ex
    With SubPath
        Dim sn As VGCore.Node
        Dim EN As VGCore.Node
        Dim q As Long
        
        If .segments.First.Type = cdrLineSegment Then
            Set sn = .AddNodeAt(dist * 2, cdrAbsoluteSegmentOffset)
            For q = 1 To sn.index
                .Nodes.First.Delete
            Next
        Else
            Set sn = .AddNodeAt(dist, cdrAbsoluteSegmentOffset)
            Set EN = .AddNodeAt(.Length - dist, cdrAbsoluteSegmentOffset)
            For q = 1 To .Nodes.count - EN.index - 1
                .Nodes.Last.Delete
            Next
    
            For q = 1 To sn.index
                .Nodes.First.Delete
            Next
        End If
    End With
ex:
End Sub
