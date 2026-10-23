Private Sub AddLinkedToActiveSelection()
    Dim q As Long
    For q = 1 To VGCore.ActiveSelectionRange.count
        VGCore.ActiveSelectionRange(q).GetLinkedShapes(cdrLinkAllConnections).AddToSelection
    Next
End Sub
