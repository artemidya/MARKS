Private Function IsDriveFixed(ByVal FilePath As String) As Boolean
    If FilePath = "" Then Exit Function
    On Error GoTo 10
    
    IsDriveFixed = True
    Dim DriveLetter As String
    DriveLetter = UCase(left(FilePath, 3)) ' например: "C:\"
    
    Dim driveType As Long
    driveType = GetDriveType(DriveLetter)
    
    IsDriveFixed = (driveType = DRIVE_FIXED)
10
End Function
