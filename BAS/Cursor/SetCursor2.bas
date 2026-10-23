Public Sub SetCursor2(CursorType As CursorTypes)
    On Error GoTo 10
    If Not IsCursorType(CursorType) Then
        SetCursor LoadCursor(0, CursorType)
        Sleep 200 ' wait a bit, needed for rendering
    End If
10
End Sub
