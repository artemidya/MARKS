Private Sub DrawCurves(ByRef Curves() As curve)
    Dim q As Long
    Dim crv As New VGCore.curve
    Dim shape As VGCore.shape
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    
    For q = CountCurves + 1 To CountLineQ
        With Curves(q).SubPath(1)
            crv.CreateSubPath(.PosX(1), .PosY(1)).AppendLineSegment .PosX(2), .PosY(2)
        End With
    Next
    'Чертит прямоугольник
    crv.AppendSubpathRectangle XSel, YSel, XSel + WithSel, YSel - HeightSel
    
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("LabirintCreator.BeginCommandGroup")
    Set shape = ActiveLayer.CreateCurve(crv)
    shape.Outline.SetProperties Color:=CreateCMYKColor(0, 60, 100, 0)
    
    sr.AddToSelection
    ActiveSelection.Group
    
    Application.ActiveDocument.EndCommandGroup
End Sub
