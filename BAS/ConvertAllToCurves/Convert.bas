Public Sub Convert()
    On Error GoTo er:
    UserData.Load
    Dim RegData1 As New RegData
    
    Application.ActiveDocument.PreserveSelection = True
    Application.Optimization = True


    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("ConvertAllToCurves.Command")
    AddLinkedToActiveSelection
    Dim OrigSelection As VGCore.ShapeRange
    Set OrigSelection = ActiveSelectionRangeByOrder()


    Dim convertedShapes As VGCore.ShapeRange
    Set convertedShapes = ConvertShapes(OrigSelection.shapes, RegData1.DelFill, RegData1.DelOutline)

    DeleteUnvisibleLines convertedShapes
er:
    On Error Resume Next
    Application.ActiveDocument.ClearSelection
    OrigSelection.Delete
    If RegData1.UngroupAll Then
        Set convertedShapes = convertedShapes.UngroupAllEx
    End If
    convertedShapes.AddToSelection

    Application.ActiveDocument.EndCommandGroup
    Application.Optimization = False
    Application.ActiveDocument.PreserveSelection = False
    Application.Refresh
End Sub
