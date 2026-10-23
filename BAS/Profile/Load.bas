Public Sub Load()
    Dim content As String
    Dim testVersion As Long
    content = ReadFile(lFile)
    testVersion = StringToLong(GetTag("MinimalVersion", content))
    If content = "" Or testVersion > UserData.PlotCalcVersion Then
        If testVersion > UserData.PlotCalcVersion Then
            MsgBox localization.ActiveLng.Item("Profile.MsgBox_1") & vbNewLine & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_2") & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_3") & vbNewLine & localization.ActiveLng.Item("Profile.MsgBox_4"), vbInformation
        End If
        SetDefault
        Save
    Else
'        Take all data from profile file
        lDevice = GetTag("Device", content)
        lOutPut = GetTag("OutPut", content)
        lName = GetTag("Name", content)
        lCommandLanguage = StringToCommandLanguage(GetTag("CommandLanguage", content))
        lMaxSheetWidth = StringToLong(GetTag("MaxSheetWidth", content))
        lMaxSheetHeight = StringToLong(GetTag("MaxSheetHeight", content))
        lVisible = StringToBoolean(GetTag("Visible", content))
        lAccelerations() = StringToLongArray(GetTag("Accelerations", content))
        lActiveAcceleration = StringToLong(GetTag("ActiveAcceleration", content))
        lSeparateParametrs = StringToBoolean(GetTag("SeparateParametrs", content))
        lOneCommandOnPath = StringToBoolean(GetTag("OneCommandOnPath", content))
        lDoublePass = StringToBoolean(GetTag("DoublePass", content))
        lSpeeds() = StringToLongArray(GetTag("Speeds", content))
        lActiveSpeed = StringToLong(GetTag("ActiveSpeed", content))
        lActiveSpeedUP = StringToLong(GetTag("ActiveSpeedUP", content))
        lForces() = StringToLongArray(GetTag("Forces", content))
        lActiveForce = StringToLong(GetTag("ActiveForce", content))
        lInitializer = GetTag("Initializer", content)
        lBeginWorkFile = GetTag("BeginWorkFile", content)
        lEndWorkFile = GetTag("EndWorkFile", content)
        lTerminator = GetTag("Terminator", content)
        lFilePLT = GetTag("FilePLT", content)
        lMemorySize = StringToLong(GetTag("MemorySize", content))
        lKnifeOffset = StringToLong(GetTag("KnifeOffset", content))
        If lKnifeOffset > 5550 Then lKnifeOffset = 5550
        If lKnifeOffset < 0 Then lKnifeOffset = 0
        lKnifeOffset = (lKnifeOffset \ 25) * 25
        lClosingCurve = StringToLong(GetTag("ClosingCurve", content))
        lClosingCurve = (lClosingCurve \ 25) * 25
        If lClosingCurve > 10000 Then lClosingCurve = 10000
        If lClosingCurve < 0 Then lClosingCurve = 0
        lCornerPoints = StringToBoolean(GetTag("CornerPoints", content))
        lZeroPosition = StringToLong(GetTag("ZeroPosition", content))
        If lZeroPosition > 2 Or lZeroPosition < 0 Then lZeroPosition = RightP
        lHeightScale = ValPD(GetTag("HeightScale", content))
        lWidthScale = ValPD(GetTag("WidthScale", content))
        lTurn = StringToLong(GetTag("Turn", content))
        lLineType = StringToLong(GetTag("LineType", content))
        lApproximationLevel = StringToLong(GetTag("ApproximationLevel", content))
        lLaser = StringToBoolean(GetTag("Laser", content))
        lLaserDX = StringToLong(GetTag("LaserDX", content))
        lLaserDY = StringToLong(GetTag("LaserDY", content))
        lForModeratorDevice = GetTag("ForModeratorDevice", content)
        lPlotterModelString = GetTag("PlotterModelString", content)
        lPlotterName = GetTag("PlotterName", content)
        lMinimalVersion = StringToLong(GetTag("MinimalVersion", content))
        lCoordinateSystem = StringToCoordinateSystem(GetTag("CoordinateSystem", content))
        NeedSave = False
    End If
End Sub
