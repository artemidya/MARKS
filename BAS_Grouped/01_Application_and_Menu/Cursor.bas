' =============================================================================
' Group: 01_Application_and_Menu
' Source module: Cursor.bas
' Procedures: 2
' =============================================================================

' -----------------------------------------------------------------------------
' SetCursor2 (Sub), source lines 46-53
' -----------------------------------------------------------------------------
Public Sub SetCursor2(CursorType As CursorTypes)
    On Error GoTo 10
    If Not IsCursorType(CursorType) Then
        SetCursor LoadCursor(0, CursorType)
        Sleep 200 ' wait a bit, needed for rendering
    End If
10
End Sub

' -----------------------------------------------------------------------------
' IsCursorType (Function), source lines 56-67
' -----------------------------------------------------------------------------
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
