Private Sub ToGroup(ShapesArray() As VGCore.shape, StartFromQ As Long)
    Dim q As Long
    Dim x As Double, y As Double
    
    Dim ForGropingSR As New VGCore.ShapeRange
    
    Dim ControlCurve As VGCore.curve
    Set ControlCurve = CreatePolylineCurveFromShape(ShapesArray(StartFromQ))
       
    Dim left As Double, right As Double
    Dim top As Double, bottom As Double
    
    With ControlCurve.boundingBox
        left = .left
        right = .right
        top = .top
        bottom = .bottom
    End With
    
    q = StartFromQ + 1
    Do While q <= UBound(ShapesArray())
        With ShapesArray(q).boundingBox
            x = (.right + .left) / 2
            If x < right And x > left Then
                y = (.top + .bottom) / 2
                If y < top And y > bottom Then
                    If IsShapeOnPolyLineCurve(ControlCurve, x, y, ShapesArray(q)) Then
'                       inside shape!
                        ForGropingSR.Add ShapesArray(q)
                        DeleteShapeFromShapesArray ShapesArray(), q
                        q = q - 1
                    End If
                End If
            End If
        End With
        q = q + 1
    Loop
    
    If ForGropingSR.count > 0 Then
        ForGropingSR.Add ShapesArray(StartFromQ)
        Set ShapesArray(StartFromQ) = ForGropingSR.Group
    End If
    
End Sub
