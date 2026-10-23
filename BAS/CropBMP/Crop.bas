Public Sub Crop()
    On Error GoTo exsub
    UserData.Load
    Application.ActiveDocument.PreserveSelection = True
    Dim BMPShapesRange As VGCore.ShapeRange
    Set BMPShapesRange = GetBMPShapeRange(VGCore.ActiveSelection.shapes)

    If BMPShapesRange.count > 0 Then
        Dim generalCurve As VGCore.curve
        Set generalCurve = GetGeneralCurve(VGCore.ActiveSelection.shapes)
        
        If generalCurve.Subpaths.count > 0 Then
            On Error GoTo ex2
            'Start
            Dim RegData1 As New RegData
            Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("CropBMP.Command") & " (" & CDbl(RegData1.CropBMPoffset) / 1000 & " " & localization.ActiveLng.Item("RaskroyForm.mm") & ")"
            Dim lastUnit As cdrUnit
            With Application.ActiveDocument
                lastUnit = .Unit
                .Unit = cdrMillimeter
                
                Dim offset As Double
                offset = RegData1.CropBMPoffset / 1000
                Set generalCurve = generalCurve.Contour(offset, cdrContourOutside, , cdrContourCornerRound)
                .Unit = lastUnit
            End With
            Application.Optimization = True
            
            'Group
            Dim g As VGCore.shape
            Set g = VGCore.ActiveSelection.shapes.All.Group
            
            Dim s As VGCore.shape
            Set s = VGCore.ActiveLayer.CreateCurve(generalCurve)

            
            Dim q As Long
            For q = 1 To BMPShapesRange.count
                Dim s2 As VGCore.shape
                Set s2 = s.Intersect(BMPShapesRange(q), True, True)
                If Not s2 Is Nothing Then
                     s2.OrderBackOf BMPShapesRange(q)
                End If
               
                BMPShapesRange(q).Delete
            Next
            s.Delete
            Application.ActiveDocument.ClearSelection
            g.AddToSelection
ex2:
            Application.ActiveDocument.EndCommandGroup
            'end
        End If
    End If
exsub:
    
    Application.Optimization = False
    Application.ActiveDocument.PreserveSelection = False
    Application.Refresh
End Sub
