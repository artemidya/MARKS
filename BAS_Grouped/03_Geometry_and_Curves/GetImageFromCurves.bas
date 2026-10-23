' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: GetImageFromCurves.frm
' Procedures: 7
' =============================================================================

' -----------------------------------------------------------------------------
' UserForm_Initialize (Sub), source lines 13-16
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    If ActiveSelection.shapes.count < 1 Then MsgBox "Select shape", vbInformation: End
    GetCurves Application.ActiveSelectionRange.shapes
End Sub

' -----------------------------------------------------------------------------
' GetCurves (Sub), source lines 17-34
' -----------------------------------------------------------------------------
Public Sub GetCurves(shapes As VGCore.shapes)
    On Error GoTo 10
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrMillimeter

       
    TextBox1 = "Private Sub InitCurves()" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Dim myCurves As New Collection" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Set xxxCol = myCurves" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Dim path As waPath" & vbNewLine
    
    GetCurvesCycle shapes

    TextBox1 = TextBox1 & "End Sub" & vbNewLine
10
    Application.ActiveDocument.Unit = lastUnit
End Sub

' -----------------------------------------------------------------------------
' GetCurvesCycle (Sub), source lines 35-55
' -----------------------------------------------------------------------------
Private Sub GetCurvesCycle(shapes As VGCore.shapes)
    Dim q1 As Long
    For q1 = shapes.count To 1 Step -1
        DoEvents
        With shapes(q1)
            If .Type = cdrCurveShape Then
            
            
                GetCurve .curve, forBrushString(shapes(q1))
                
            ElseIf .Type = cdrGroupShape Or .Type = cdrContourGroupShape Or .Type = cdrArtisticMediaGroupShape Then
                GetCurvesCycle .shapes
            ElseIf TryDisplayCurve(shapes(q1)) Then
                If .Type <> cdrBitmapShape Then GetCurve .DisplayCurve, forBrushString(shapes(q1))
            Else
                'for text
                TryConvertToCurves shapes(q1)
            End If
        End With
    Next
End Sub

' -----------------------------------------------------------------------------
' forBrushString (Function), source lines 56-103
' -----------------------------------------------------------------------------
Private Function forBrushString(shape As VGCore.shape) As String
    Dim c As Long:    c = 1000
    
    Dim R As Long
    Dim g As Long
    Dim b As Long
    
    Dim R2 As Long
    Dim G2 As Long
    Dim b2 As Long
    
    Dim x1 As Long
    Dim y1 As Long
    Dim x2 As Long
    Dim y2 As Long
    
    forBrushString = vbTab & vbTab & "Set .Brush = new waBrush" & vbNewLine
    
    With shape
        If .Fill.Type = cdrUniformFill Then
            .Fill.UniformColor.ConvertToRGB
            R = .Fill.UniformColor.RGBRed
            g = .Fill.UniformColor.RGBGreen
            b = .Fill.UniformColor.RGBBlue
            forBrushString = forBrushString & vbTab & vbTab & ".Brush.SetColor RGB(" & R & ", " & g & ", " & b & "), 100"

        ElseIf .Fill.Type = cdrFountainFill Then
            .Fill.Fountain.colors.First.Color.ConvertToRGB
            .Fill.Fountain.colors.Last.Color.ConvertToRGB
            
            R = .Fill.Fountain.colors.First.Color.RGBRed
            g = .Fill.Fountain.colors.First.Color.RGBGreen
            b = .Fill.Fountain.colors.First.Color.RGBBlue
            
            R2 = .Fill.Fountain.colors.Last.Color.RGBRed
            G2 = .Fill.Fountain.colors.Last.Color.RGBGreen
            b2 = .Fill.Fountain.colors.Last.Color.RGBBlue
            
            x1 = .Fill.Fountain.startX
            y1 = .Fill.Fountain.startY
            x2 = .Fill.Fountain.EndX
            y2 = .Fill.Fountain.EndY
            
            '.Brush.SetLinearGradient x1, y1, x2, y2, RGB(r, g, b), RGB(r2, g2, b2), WrapModeTileFlipXY
            forBrushString = forBrushString & vbTab & vbTab & ".Brush.SetLinearGradient " & x1 * c & ", " & y1 * c & ", " & x2 * c & ", " & y2 * c & ", RGB(" & R & ", " & g & ", " & b & "), RGB(" & R2 & ", " & G2 & ", " & b2 & "), WrapModeTileFlipXY"
        End If
    End With
End Function

' -----------------------------------------------------------------------------
' TryConvertToCurves (Sub), source lines 104-118
' -----------------------------------------------------------------------------
Private Sub TryConvertToCurves(shape As VGCore.shape)
    Application.Optimization = True
    Dim s2 As VGCore.shape
    Set s2 = shape.Clone
    On Error GoTo 10
    s2.ConvertToCurves
    If s2.Type = cdrCurveShape Then
        GetCurve s2.curve, forBrushString(s2)
    ElseIf s2.Type = cdrGroupShape Then
        GetCurvesCycle s2.shapes
    End If
10
    s2.Delete
    Application.Optimization = False
End Sub

' -----------------------------------------------------------------------------
' TryDisplayCurve (Function), source lines 119-123
' -----------------------------------------------------------------------------
Private Function TryDisplayCurve(shape As VGCore.shape) As Boolean
    On Error GoTo 10
    If shape.DisplayCurve.Nodes.count > 0 Then TryDisplayCurve = True
10
End Function

' -----------------------------------------------------------------------------
' GetCurve (Sub), source lines 124-156
' -----------------------------------------------------------------------------
Private Sub GetCurve(crv As VGCore.curve, forBrushString As String)
    Dim c As Long:    c = 1000
    With crv
                
        TextBox1 = TextBox1 & vbNewLine & vbTab & "Set path = New waPath" & vbNewLine
        TextBox1 = TextBox1 & vbTab & "With path" & vbNewLine
        Dim q As Long, q2 As Long
        For q = 1 To .Subpaths.count
            With .Subpaths(q)

                TextBox1 = TextBox1 & vbTab & vbTab & ".StartSubPath" & vbNewLine
                With .StartNode
                    TextBox1 = TextBox1 & vbTab & vbTab & ".AddLineSegment " & CLng(.PositionX * c) & ", " & CLng(.PositionY * c) & vbNewLine
                End With
                For q2 = 1 To .segments.count
                    With .segments(q2)
                        Select Case .Type
                            Case cdrCurveSegment
                                TextBox1 = TextBox1 & vbTab & vbTab & ".AddBezierSegment " & CLng(.StartNode.PositionX * c) & ", " & CLng(.StartNode.PositionY * c) & ", " & CLng(.StartingControlPointX * c) & ", " & CLng(.StartingControlPointY * c) & ", " & CLng(.EndingControlPointX * c) & ", " & CLng(.EndingControlPointY * c) & ", " & CLng(.EndNode.PositionX * c) & ", " & CLng(.EndNode.PositionY * c) & vbNewLine
                            Case cdrLineSegment
                                TextBox1 = TextBox1 & vbTab & vbTab & ".AddLineSegment " & CLng(.EndNode.PositionX * c) & ", " & CLng(.EndNode.PositionY * c) & vbNewLine
                        End Select
                    End With
                Next
                If .closed Then TextBox1 = TextBox1 & vbTab & vbTab & ".CloseSubPath" & vbNewLine & vbNewLine

            End With
        Next
        TextBox1 = TextBox1 & forBrushString & vbNewLine
        TextBox1 = TextBox1 & vbTab & "End With" & vbNewLine
        TextBox1 = TextBox1 & vbTab & "myCurves.Add path" & vbNewLine
    End With
End Sub
