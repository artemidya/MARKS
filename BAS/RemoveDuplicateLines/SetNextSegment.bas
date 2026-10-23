Private Sub SetNextSegment(Node As VGCore.Node, ByRef BreakSegment As VGCore.segment)
    On Error GoTo 10
    Set BreakSegment = Node.NextSegment
10
End Sub
