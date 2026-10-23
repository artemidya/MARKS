Private Sub CreateAfterCutMark()
    Application.ActiveDocument.BeginCommandGroup "Mark"
    '\ /
    ' X
    '/ \
    Dim crv As VGCore.curve
    Dim P1 As waGraphicsShared.POINTF
    Dim P2 As waGraphicsShared.POINTF
    Dim P3 As waGraphicsShared.POINTF
    Dim P4 As waGraphicsShared.POINTF
    Dim d As Double:  d = 0.5
    
    With VGCore.ActiveSelection
        P1.x = .LeftX - d
        P1.y = .BottomY - d
        P2.x = .RightX + d
        P2.y = .TopY + d
        
        P3.x = .LeftX - d
        P3.y = .TopY + d
        P4.x = .RightX + d
        P4.y = .BottomY - d
    End With
    
    Set crv = VGCore.ActiveDocument.CreateCurve
    With crv.CreateSubPath(P1.x, P1.y)
        .AppendLineSegment P2.x, P2.y
    End With

    With crv.CreateSubPath(P3.x, P3.y)
        .AppendLineSegment P4.x, P4.y
    End With
    
    Dim s1 As VGCore.shape
    Set s1 = VGCore.ActiveLayer.CreateCurve(crv)
    s1.Outline.SetProperties Width:=0.1, Color:=CreateCMYKColor(0, 100, 100, 0)
    
    Application.ActiveDocument.EndCommandGroup
End Sub
