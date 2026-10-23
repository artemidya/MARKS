Public Sub TrimCrv()
    On Error GoTo er:
    UserData.Load
    Application.ActiveDocument.PreserveSelection = True
    Application.Optimization = True

    Dim q As Long
    Set GdelShapeRange = New VGCore.ShapeRange
    Set GorderBackShapeRange = New VGCore.ShapeRange
    
    Dim lastUnit As cdrUnit
    With Application.ActiveDocument
        lastUnit = .Unit
        .Unit = cdrMillimeter
    End With
    
    
    
    
    Dim RegData1 As New RegData
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("TrimCurves.Command") & " (" & CDbl(RegData1.TrimOffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
    BreakDistance = RegData1.TrimOffset / 1000
    If BreakDistance = 0 Then BreakDistance = -1
    
    Dim OrigSelection As VGCore.ShapeRange
    Set OrigSelection = VGCore.ActiveSelectionRange
    TrimShapes OrigSelection.shapes
    
    
    
    On Error Resume Next
    For q = 1 To GorderBackShapeRange.count
        If GdelShapeRange(q).ParentGroup Is Nothing Then
            GorderBackShapeRange(q).OrderBackOf GdelShapeRange(q)
        ElseIf GdelShapeRange(q).ParentGroup.Type = cdrGroupShape Then
            GorderBackShapeRange(q).OrderBackOf GdelShapeRange(q)
        End If
    Next
    
    
er:
    'Kill shapes
    GdelShapeRange.All.Delete
    
    Application.ActiveDocument.ClearSelection
    OrigSelection.All.AddToSelection
    For q = 1 To GorderBackShapeRange.count
        If GorderBackShapeRange(q).ParentGroup Is Nothing Then
            GorderBackShapeRange(q).AddToSelection
        End If
    Next
    Application.ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.ActiveDocument.PreserveSelection = False
    Application.Refresh
    
    Application.ActiveDocument.Unit = lastUnit
End Sub
