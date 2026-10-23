Private Function IsCursorType(CursorType As CursorTypes) As Boolean
    Dim CursorHandle As Long: CursorHandle = LoadCursor(ByVal 0&, CursorType)
    Dim Cursor As CursorInfo: Cursor.cbSize = Len(Cursor)
    Dim CursorInfo As Boolean: CursorInfo = GetCursorInfo(Cursor)

    If Not CursorInfo Then
        IsCursorType = False
        Exit Function
    End If

    IsCursorType = (Cursor.hCursor = CursorHandle)
End Function
