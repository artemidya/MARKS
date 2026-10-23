Public Sub SetRandomMaterialColor()
    Randomize Timer
    Dim R As Byte, g As Byte, b As Byte, a As Long
    If Int(Rnd * 10) = 1 Then
        a = Int(Rnd * 50) + 50
    Else
        a = 100
    End If
    R = Int(Rnd * 255)
    g = Int(Rnd * 255)
    b = Int(Rnd * 255)
    Set MaterialBrush = CreateBrush(RGB(R, g, b), a)
End Sub
