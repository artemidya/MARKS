Private Function Angel(segment As VGCore.segment) As Double
    Static x1 As Double
    Static y1 As Double
    Static x2 As Double
    Static y2 As Double
    
    With segment
        With .StartNode
            x1 = .PositionX
            y1 = .PositionY
        End With
        With .EndNode
            x2 = .PositionX
            y2 = .PositionY
        End With
    End With

    If x2 <> x1 Then
        Angel = Atn((y2 - y1) / (x2 - x1))
    ElseIf y2 <> y1 Then
        Angel = 1.5707963267949
    Else
        Err.Raise vbObjectError, , ""
    End If
End Function
