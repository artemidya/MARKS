Private Sub NoFill(ByRef s1 As ShapeRange)
    On Error Resume Next
    s1.ApplyNoFill
    s1.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, 100)
End Sub
