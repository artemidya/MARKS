Public Sub StartReGroup()
    If Processing Then Exit Sub
    UserData.Load
    If VGCore.ActiveSelectionRange.count > 0 Then
        On Error GoTo er:
        Dim RegData1 As New RegData

        Processing = True
        ProgressFormIsShowed = False
        Application.Optimization = True
        Application.ActiveDocument.PreserveSelection = True
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("ReGroup.Command")
        
        Dim sr As New VGCore.ShapeRange
        sr.AddRange VGCore.ActiveSelectionRange.UngroupAllEx
        
        StartTime = Timer
        LastRefreshTime = Timer
        
        If RegData1.ConnectingCurves Then
            '1. Connecting
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ConnectingLines")
            ProgressForm.EnableSkip = True
            ConnectingCurves sr
            If ProgressForm.Break Then GoTo er
            ProgressForm.EnableSkip = False
            ProgressForm.Percent = 1
    
            '2. Connecting subpaths
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ConnectingSubpaths")
            SelfyConnectingShapes sr
        End If
        
        '3. Sorting
        ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.Sorting")
        sr.Sort " @shape1.Width * @shape1.Height > @shape2.Width * @shape2.Height"
        
        '4. Add to ShapesArrey
        Dim ShapesArray() As VGCore.shape
        ReDim ShapesArray(sr.count - 1)
        Dim q As Long
        For q = 1 To sr.count
            Set ShapesArray(q - 1) = sr(q)
        Next
        
        '5. ReGrouping
        ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.Command")
        ProgressForm.Percent = 0
        StartR ShapesArray()
        If ProgressForm.Break Then GoTo er
        ProgressForm.Percent = 1
        
        
        Set sr = New VGCore.ShapeRange
        For q = 0 To UBound(ShapesArray())
            sr.Add ShapesArray(q)
        Next
    
        If RegData1.ReduceCurves Then
            '6. Reduce
            ProgressForm.EnableSkip = True
            ProgressForm.StageName = localization.ActiveLng.Item("ReGroup.ReduceNodes")
            CountShapes = 0
            ReadyCountShapes = 0
            GetCountCurves sr.shapes
            Reduce sr.shapes
            ProgressForm.EnableSkip = False
        End If
        
        Application.ActiveDocument.ClearSelection
        sr.All.AddToSelection
        
er:
        ProgressForm.MbClosing = True
        Application.ActiveDocument.EndCommandGroup
        On Error Resume Next
        If ProgressForm.Break Then
            Application.ActiveDocument.Undo
        End If
        Unload ProgressForm
        
        Application.Optimization = False
        Application.ActiveDocument.PreserveSelection = False
        Processing = False
        Application.Refresh
    End If
End Sub
