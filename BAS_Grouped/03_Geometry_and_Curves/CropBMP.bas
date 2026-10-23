' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: CropBMP.bas
' Procedures: 6
' =============================================================================

' -----------------------------------------------------------------------------
' Crop (Sub), source lines 4-63
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' GetBMPShapeRange (Function), source lines 64-75
' -----------------------------------------------------------------------------
Private Function GetBMPShapeRange(shapes As VGCore.shapes) As VGCore.ShapeRange
    Set GetBMPShapeRange = Application.CreateShapeRange
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrGroupShape
                GetBMPShapeRange.AddRange GetBMPShapeRange(shapes(q).shapes)
            Case cdrShapeType.cdrBitmapShape
                GetBMPShapeRange.Add shapes(q)
         End Select
    Next
End Function

' -----------------------------------------------------------------------------
' GetGeneralCurve (Function), source lines 77-99
' -----------------------------------------------------------------------------
Private Function GetGeneralCurve(shapes As VGCore.shapes) As VGCore.curve
    Set GetGeneralCurve = Application.CreateCurve
    Dim q As Long
    For q = 1 To shapes.count
         Select Case shapes(q).Type
            Case cdrShapeType.cdrCurveShape
                Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(shapes(q).curve))
            Case cdrShapeType.cdrGroupShape, cdrShapeType.cdrBlendGroupShape, cdrShapeType.cdrContourGroupShape, cdrArtisticMediaGroupShape
                Set GetGeneralCurve = GetGeneralCurve.WeldWith(GetGeneralCurve(shapes(q).shapes))
            Case cdrShapeType.cdrBitmapShape
                'Do Nothing
            Case Else
                If TryDisplayCurve(shapes(q)) Then
                    Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(shapes(q).DisplayCurve))
                Else
                    Dim NewCrv As VGCore.curve
                    If TryConvertToCurves(shapes(q), NewCrv) Then
                        Set GetGeneralCurve = GetGeneralCurve.WeldWith(CloseCurve(NewCrv))
                    End If
                End If
         End Select
    Next
End Function

' -----------------------------------------------------------------------------
' TryConvertToCurves (Function), source lines 100-112
' -----------------------------------------------------------------------------
Private Function TryConvertToCurves(shape As VGCore.shape, NewCrv As VGCore.curve) As Boolean
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        TryConvertToCurves = True
        Set NewCrv = Application.CreateCurve
        NewCrv.AppendCurve s2.curve
    End If
10
    s2.Delete
End Function

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 113-120
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As shape) As Boolean
    On Error GoTo Err
    Dim c As Long
    c = shape.DisplayCurve.Subpaths.count
    TryDisplayCurve = True
    Exit Function
Err:
End Function

' -----------------------------------------------------------------------------
' CloseCurve (Function), source lines 121-129
' -----------------------------------------------------------------------------
Private Function CloseCurve(crv As VGCore.curve) As VGCore.curve
    If crv.closed Then
        Set CloseCurve = crv
    Else
        Set CloseCurve = VGCore.ActiveDocument.CreateCurve
        CloseCurve.AppendCurve crv
        CloseCurve.closed = True
    End If
End Function
