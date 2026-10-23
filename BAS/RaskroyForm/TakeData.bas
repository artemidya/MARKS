Private Sub TakeData()
    On Error GoTo Err:
    Dim q1 As Long
    Dim qq As Long
    Dim CurveFromBitmap As curve

    With VGCore.ActiveSelection
        ReDim RCurves(1 To .shapes.count)
        ReDim Du(1 To .shapes.count)
        ReDim RerotateCrv(1 To .shapes.count)
        For q1 = 1 To .shapes.count
            ProgressBar1.value = Int(q1 / .shapes.count * 10000)
            If Break = True Then Exit Sub
            DoEvents
    
            If Not TestShapeForUsedShapes(.shapes(q1).StaticId, Du(), qq) Then
                qq = qq + 1
                Set Du(qq) = .shapes(q1).GetLinkedShapes(cdrLinkAllConnections)
                If Du(qq).count = 0 Then Du(qq).Add .shapes(q1)
                Set CurveFromBitmap = VGCore.ActiveDocument.CreateCurve
                Set RerotateCrv(qq) = CreateCurveFromShapes(Du(qq).shapes, CurveFromBitmap)
                If RerotateCrv(qq).Subpaths.count = 0 Then
                     Set RerotateCrv(qq) = RerotateCrv(qq).WeldWith(CurveFromBitmap)
                End If
                RCurves(qq) = TakeRCurve(RerotateCrv(qq))
            End If
        Next
        If qq < .shapes.count Then
            ReDim Preserve RCurves(1 To qq)
            ReDim Preserve Du(1 To qq)
            ReDim Preserve RerotateCrv(1 To qq)
        End If
    End With
Err:
End Sub
