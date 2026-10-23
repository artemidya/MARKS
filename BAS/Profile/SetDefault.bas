Private Sub SetDefault()
    lName = localization.ActiveLng.Item("Profile.ProfileName") & MyProfileNumber
    lCommandLanguage = HPGL
    lMaxSheetWidth = 0
    lMaxSheetHeight = 0
    If MyProfileNumber > 4 Then
        lVisible = False
    Else
        lVisible = True
    End If
    lDevice = localization.ActiveLng.Item("Profile.Device") '"Standard device with HPGL command support"
    lInitializer = "IN;"
    lFilePLT = "*.PLT"
    lMemorySize = 4000000
    lKnifeOffset = 250
    lCornerPoints = False
    lHeightScale = 1
    lWidthScale = 1
    lApproximationLevel = Normal
    lTurn = Turn0
    lMinimalVersion = ProfileVersion
    lCoordinateSystem = Absolute
    lSeparateParametrs = False
    lOneCommandOnPath = False
    lDoublePass = False
End Sub
