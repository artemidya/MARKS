' =============================================================================
' Group: 07_Profiles_and_Settings
' Source module: Profile.cls
' Procedures: 99
' =============================================================================

' -----------------------------------------------------------------------------
' Class_Initialize (Sub), source lines 55-61
' -----------------------------------------------------------------------------
Private Sub Class_Initialize()
    CountProfiles = CountProfiles + 1
    MyProfileNumber = CountProfiles
    'Load
    lFile = VGCore.GMSManager.Projects("PlotCalc").FilePath & "Profile" & MyProfileNumber & ".xml"
    Load
End Sub

' -----------------------------------------------------------------------------
' CoordinateSystem (PropertyGet), source lines 62-64
' -----------------------------------------------------------------------------
Public Property Get CoordinateSystem() As CoordinateSystems
    CoordinateSystem = lCoordinateSystem
End Property

' -----------------------------------------------------------------------------
' CoordinateSystem (PropertyLet), source lines 65-68
' -----------------------------------------------------------------------------
Public Property Let CoordinateSystem(oValue As CoordinateSystems)
    If oValue <> lCoordinateSystem Then NeedSave = True
    lCoordinateSystem = oValue
End Property

' -----------------------------------------------------------------------------
' PlotterModelString (PropertyGet), source lines 70-72
' -----------------------------------------------------------------------------
Public Property Get PlotterModelString() As String
    PlotterModelString = lPlotterModelString
End Property

' -----------------------------------------------------------------------------
' PlotterModelString (PropertyLet), source lines 73-76
' -----------------------------------------------------------------------------
Public Property Let PlotterModelString(oValue As String)
    If oValue <> lPlotterModelString Then NeedSave = True
    lPlotterModelString = oValue
End Property

' -----------------------------------------------------------------------------
' SeparateParametrs (PropertyGet), source lines 79-81
' -----------------------------------------------------------------------------
Public Property Get SeparateParametrs() As Boolean
    SeparateParametrs = lSeparateParametrs
End Property

' -----------------------------------------------------------------------------
' SeparateParametrs (PropertyLet), source lines 82-85
' -----------------------------------------------------------------------------
Public Property Let SeparateParametrs(oValue As Boolean)
    If oValue <> lSeparateParametrs Then NeedSave = True
    lSeparateParametrs = oValue
End Property

' -----------------------------------------------------------------------------
' OneCommandOnPath (PropertyGet), source lines 86-88
' -----------------------------------------------------------------------------
Public Property Get OneCommandOnPath() As Boolean
    OneCommandOnPath = lOneCommandOnPath
End Property

' -----------------------------------------------------------------------------
' OneCommandOnPath (PropertyLet), source lines 89-92
' -----------------------------------------------------------------------------
Public Property Let OneCommandOnPath(oValue As Boolean)
    If oValue <> lOneCommandOnPath Then NeedSave = True
    lOneCommandOnPath = oValue
End Property

' -----------------------------------------------------------------------------
' DoublePass (PropertyGet), source lines 93-95
' -----------------------------------------------------------------------------
Public Property Get DoublePass() As Boolean
    DoublePass = lDoublePass
End Property

' -----------------------------------------------------------------------------
' DoublePass (PropertyLet), source lines 96-99
' -----------------------------------------------------------------------------
Public Property Let DoublePass(oValue As Boolean)
    If oValue <> lDoublePass Then NeedSave = True
    lDoublePass = oValue
End Property

' -----------------------------------------------------------------------------
' ForModeratorDevice (PropertyGet), source lines 100-102
' -----------------------------------------------------------------------------
Public Property Get ForModeratorDevice() As String
    ForModeratorDevice = lForModeratorDevice
End Property

' -----------------------------------------------------------------------------
' ForModeratorDevice (PropertyLet), source lines 103-107
' -----------------------------------------------------------------------------
Public Property Let ForModeratorDevice(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lForModeratorDevice Then NeedSave = True
    lForModeratorDevice = oValue
End Property

' -----------------------------------------------------------------------------
' content (PropertyGet), source lines 108-149
' -----------------------------------------------------------------------------
Public Property Get content() As String
    content = content & PutTag("Device", lDevice)
    content = content & PutTag("OutPut", lOutPut)
    content = content & PutTag("Name", Trim(lName))
    content = content & PutTag("CommandLanguage", CommandLanguageToString(lCommandLanguage))
    content = content & PutTag("CoordinateSystem", CoordinateSystemToString(lCoordinateSystem))
    
    content = content & PutTag("MaxSheetWidth", CStr(lMaxSheetWidth))
    content = content & PutTag("MaxSheetHeight", CStr(lMaxSheetHeight))
    content = content & PutTag("Visible", BooleanToString(lVisible))
    content = content & PutTag("Accelerations", LongArrayToString(lAccelerations))
    content = content & PutTag("ActiveAcceleration", CStr(lActiveAcceleration))
    content = content & PutTag("SeparateParametrs", BooleanToString(lSeparateParametrs))
    content = content & PutTag("OneCommandOnPath", BooleanToString(lOneCommandOnPath))
    content = content & PutTag("DoublePass", BooleanToString(lDoublePass))
    content = content & PutTag("Speeds", LongArrayToString(lSpeeds))
    content = content & PutTag("ActiveSpeed", CStr(lActiveSpeed))
    content = content & PutTag("ActiveSpeedUP", CStr(lActiveSpeedUP))
    content = content & PutTag("Forces", LongArrayToString(lForces))
    content = content & PutTag("ActiveForce", CStr(lActiveForce))
    content = content & PutTag("Initializer", lInitializer)
    content = content & PutTag("BeginWorkFile", lBeginWorkFile)
    content = content & PutTag("EndWorkFile", lEndWorkFile)
    content = content & PutTag("Terminator", lTerminator)
    content = content & PutTag("FilePLT", lFilePLT)
    content = content & PutTag("MemorySize", CStr(lMemorySize))
    content = content & PutTag("KnifeOffset", CStr(lKnifeOffset))
    content = content & PutTag("ClosingCurve", CStr(lClosingCurve))
    content = content & PutTag("CornerPoints", BooleanToString(lCornerPoints))
    content = content & PutTag("ZeroPosition", CStr(lZeroPosition))
    content = content & PutTag("HeightScale", CStr(lHeightScale))
    content = content & PutTag("WidthScale", CStr(lWidthScale))
    content = content & PutTag("Turn", CStr(lTurn))
    content = content & PutTag("LineType", CStr(lLineType))
    content = content & PutTag("ApproximationLevel", CStr(lApproximationLevel))
    content = content & PutTag("Laser", BooleanToString(lLaser))
    content = content & PutTag("LaserDX", CStr(lLaserDX))
    content = content & PutTag("LaserDY", CStr(lLaserDY))
    content = content & PutTag("ForModeratorDevice", lForModeratorDevice)
    content = content & PutTag("PlotterModelString", lPlotterModelString)
    content = content & PutTag("MinimalVersion", CStr(ProfileVersion))
End Property

' -----------------------------------------------------------------------------
' LaserDY (PropertyGet), source lines 150-152
' -----------------------------------------------------------------------------
Public Property Get LaserDY() As Long
    LaserDY = lLaserDY
End Property

' -----------------------------------------------------------------------------
' LaserDY (PropertyLet), source lines 153-158
' -----------------------------------------------------------------------------
Public Property Let LaserDY(oValue As Long)
    If oValue > 200 Then oValue = 200
    If oValue < -200 Then oValue = -200
    If oValue <> lLaserDY Then NeedSave = True
    lLaserDY = oValue
End Property

' -----------------------------------------------------------------------------
' LaserDX (PropertyGet), source lines 159-161
' -----------------------------------------------------------------------------
Public Property Get LaserDX() As Long
    LaserDX = lLaserDX
End Property

' -----------------------------------------------------------------------------
' LaserDX (PropertyLet), source lines 162-167
' -----------------------------------------------------------------------------
Public Property Let LaserDX(oValue As Long)
    If oValue > 200 Then oValue = 200
    If oValue < -200 Then oValue = -200
    If oValue <> lLaserDX Then NeedSave = True
    lLaserDX = oValue
End Property

' -----------------------------------------------------------------------------
' Laser (PropertyGet), source lines 168-170
' -----------------------------------------------------------------------------
Public Property Get Laser() As Boolean
    Laser = lLaser
End Property

' -----------------------------------------------------------------------------
' Laser (PropertyLet), source lines 171-174
' -----------------------------------------------------------------------------
Public Property Let Laser(oValue As Boolean)
    If oValue <> lLaser Then NeedSave = True
    lLaser = oValue
End Property

' -----------------------------------------------------------------------------
' MemorySize (PropertyGet), source lines 175-178
' -----------------------------------------------------------------------------
Public Property Get MemorySize() As Long
    If lMemorySize < 10000 Then MemorySize = 10000
    MemorySize = lMemorySize
End Property

' -----------------------------------------------------------------------------
' MemorySize (PropertyLet), source lines 179-183
' -----------------------------------------------------------------------------
Public Property Let MemorySize(oValue As Long)
    If oValue < 10000 Then oValue = 10000
    If oValue <> lMemorySize Then NeedSave = True
    lMemorySize = oValue
End Property

' -----------------------------------------------------------------------------
' FilePLT (PropertyGet), source lines 184-186
' -----------------------------------------------------------------------------
Public Property Get FilePLT() As String
    FilePLT = lFilePLT
End Property

' -----------------------------------------------------------------------------
' FilePLT (PropertyLet), source lines 187-191
' -----------------------------------------------------------------------------
Public Property Let FilePLT(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lFilePLT Then NeedSave = True
    lFilePLT = oValue
End Property

' -----------------------------------------------------------------------------
' ApproximationLevel (PropertyGet), source lines 192-195
' -----------------------------------------------------------------------------
Public Property Get ApproximationLevel() As ApproximationLevels
    If lApproximationLevel > 2 Or lApproximationLevel < 0 Then lApproximationLevel = Normal
    ApproximationLevel = lApproximationLevel
End Property

' -----------------------------------------------------------------------------
' ApproximationLevel (PropertyLet), source lines 196-200
' -----------------------------------------------------------------------------
Public Property Let ApproximationLevel(oValue As ApproximationLevels)
    If oValue > 2 Or oValue < 0 Then oValue = Normal
    If oValue <> lApproximationLevel Then NeedSave = True
    lApproximationLevel = oValue
End Property

' -----------------------------------------------------------------------------
' LineType (PropertyGet), source lines 201-204
' -----------------------------------------------------------------------------
Public Property Get LineType() As LineTypes
    If lLineType > 1 Or lLineType < 0 Then lLineType = lines
    LineType = lLineType
End Property

' -----------------------------------------------------------------------------
' LineType (PropertyLet), source lines 205-209
' -----------------------------------------------------------------------------
Public Property Let LineType(oValue As LineTypes)
    If oValue > 1 Or oValue < 0 Then oValue = lines
    If oValue <> lLineType Then NeedSave = True
    lLineType = oValue
End Property

' -----------------------------------------------------------------------------
' Turn (PropertyGet), source lines 210-214
' -----------------------------------------------------------------------------
Public Property Get Turn() As Turns
    If lTurn > 3 Then lTurn = Turn0
    If lTurn < 0 Then lTurn = Turn0
    Turn = lTurn
End Property

' -----------------------------------------------------------------------------
' Turn (PropertyLet), source lines 215-218
' -----------------------------------------------------------------------------
Public Property Let Turn(oValue As Turns)
    If oValue <> lTurn Then NeedSave = True
    lTurn = oValue
End Property

' -----------------------------------------------------------------------------
' HeightScale (PropertyGet), source lines 219-225
' -----------------------------------------------------------------------------
Public Property Get HeightScale() As Double
    If lHeightScale > 5 Then lHeightScale = 5
    If lHeightScale < -5 Then lHeightScale = -5
    If lHeightScale >= 0 And lHeightScale < 0.5 Then lHeightScale = 0.5
    If lHeightScale > -0.5 And lHeightScale < 0 Then lHeightScale = -0.5
    HeightScale = lHeightScale
End Property

' -----------------------------------------------------------------------------
' HeightScale (PropertyLet), source lines 226-232
' -----------------------------------------------------------------------------
Public Property Let HeightScale(oValue As Double)
    If oValue = 0 Then oValue = 1
    If oValue > 5 Then oValue = 5
    If oValue < -5 Then oValue = -5
    If oValue <> lHeightScale Then NeedSave = True
    lHeightScale = oValue
End Property

' -----------------------------------------------------------------------------
' HeightMirror (PropertyGet), source lines 233-235
' -----------------------------------------------------------------------------
Public Property Get HeightMirror() As Boolean
    If lHeightScale < 0 Then HeightMirror = True
End Property

' -----------------------------------------------------------------------------
' WidthMirror (PropertyGet), source lines 236-238
' -----------------------------------------------------------------------------
Public Property Get WidthMirror() As Boolean
    If lWidthScale < 0 Then WidthMirror = True
End Property

' -----------------------------------------------------------------------------
' WidthScale (PropertyGet), source lines 239-245
' -----------------------------------------------------------------------------
Public Property Get WidthScale() As Double
    If lWidthScale > 5 Then lWidthScale = 5
    If lWidthScale < -5 Then lWidthScale = -5
    If lWidthScale >= 0 And lWidthScale < 0.5 Then lWidthScale = 0.5
    If lWidthScale > -0.5 And lWidthScale < 0 Then lWidthScale = -0.5
    WidthScale = lWidthScale
End Property

' -----------------------------------------------------------------------------
' WidthScale (PropertyLet), source lines 246-252
' -----------------------------------------------------------------------------
Public Property Let WidthScale(oValue As Double)
    If oValue = 0 Then oValue = 1
    If oValue > 5 Then oValue = 5
    If oValue < -5 Then oValue = -5
    If oValue <> lWidthScale Then NeedSave = True
    lWidthScale = oValue
End Property

' -----------------------------------------------------------------------------
' ZeroPosition (PropertyGet), source lines 253-256
' -----------------------------------------------------------------------------
Public Property Get ZeroPosition() As ZeroPositions
    If lZeroPosition > 2 Or lZeroPosition < 0 Then lZeroPosition = RightP
    ZeroPosition = lZeroPosition
End Property

' -----------------------------------------------------------------------------
' ZeroPosition (PropertyLet), source lines 257-261
' -----------------------------------------------------------------------------
Public Property Let ZeroPosition(oValue As ZeroPositions)
    If oValue > 2 Or oValue < 0 Then oValue = 0
    If oValue <> lZeroPosition Then NeedSave = True
    lZeroPosition = oValue
End Property

' -----------------------------------------------------------------------------
' CornerPoints (PropertyGet), source lines 262-264
' -----------------------------------------------------------------------------
Public Property Get CornerPoints() As Boolean
    CornerPoints = lCornerPoints
End Property

' -----------------------------------------------------------------------------
' CornerPoints (PropertyLet), source lines 265-268
' -----------------------------------------------------------------------------
Public Property Let CornerPoints(oValue As Boolean)
    If oValue <> lCornerPoints Then NeedSave = True
    lCornerPoints = oValue
End Property

' -----------------------------------------------------------------------------
' KnifeOffset (PropertyGet), source lines 269-271
' -----------------------------------------------------------------------------
Public Property Get KnifeOffset() As Long
    KnifeOffset = lKnifeOffset
End Property

' -----------------------------------------------------------------------------
' KnifeOffset (PropertyLet), source lines 272-278
' -----------------------------------------------------------------------------
Public Property Let KnifeOffset(oValue As Long)
    If oValue > 5550 Then oValue = 5550
    If oValue < 0 Then oValue = 0
    oValue = (oValue \ 25) * 25
    If oValue <> lKnifeOffset Then NeedSave = True
    lKnifeOffset = oValue
End Property

' -----------------------------------------------------------------------------
' ClosingCurve (PropertyGet), source lines 279-281
' -----------------------------------------------------------------------------
Public Property Get ClosingCurve() As Long
    ClosingCurve = lClosingCurve
End Property

' -----------------------------------------------------------------------------
' ClosingCurve (PropertyLet), source lines 282-288
' -----------------------------------------------------------------------------
Public Property Let ClosingCurve(oValue As Long)
    oValue = (oValue \ 25) * 25
    If oValue > 10000 Then oValue = 10000
    If oValue < 0 Then oValue = 0
    If oValue <> lClosingCurve Then NeedSave = True
    lClosingCurve = oValue
End Property

' -----------------------------------------------------------------------------
' Load (Sub), source lines 289-351
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' Save (Sub), source lines 352-354
' -----------------------------------------------------------------------------
Public Sub Save()
    If SaveFile(lFile, Me.content) Then NeedSave = False
End Sub

' -----------------------------------------------------------------------------
' SetDefault (Sub), source lines 355-380
' -----------------------------------------------------------------------------
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

' -----------------------------------------------------------------------------
' ReadFile (Function), source lines 381-388
' -----------------------------------------------------------------------------
Private Function ReadFile(fileName As String) As String
      On Error GoTo 10
      Dim MyFile As Integer:  MyFile = FreeFile
      Open fileName For Input As #MyFile
            ReadFile = Input$(LOF(MyFile), #MyFile)
      Close #MyFile
10
End Function

' -----------------------------------------------------------------------------
' SaveFile (Function), source lines 389-397
' -----------------------------------------------------------------------------
Private Function SaveFile(fileName As String, ByRef content As String) As Boolean
      On Error GoTo 10
      Dim MyFile As Integer:  MyFile = FreeFile
      Open fileName For Output As #MyFile
            Print #MyFile, content
      Close #MyFile
      SaveFile = True
10
End Function

' -----------------------------------------------------------------------------
' PutTag (Function), source lines 398-400
' -----------------------------------------------------------------------------
Private Function PutTag(TagName As String, ByRef text As String) As String
      PutTag = "<" & TagName & ">" & text & "</" & TagName & ">" & vbNewLine
End Function

' -----------------------------------------------------------------------------
' GetTag (Function), source lines 401-410
' -----------------------------------------------------------------------------
Private Function GetTag(TagName As String, ByRef FromText As String) As String
      Dim s As String, e As String
      s = InStr(1, FromText, "<" & TagName & ">", vbTextCompare)
      If s <> 0 Then
            e = InStr(s + Len(TagName) + 2, FromText, "</" & TagName & ">", vbTextCompare)
            If e <> 0 Then
                  GetTag = Mid(FromText, s + Len(TagName) + 2, e - s - Len(TagName) - 2)
            End If
      End If
End Function

' -----------------------------------------------------------------------------
' OutPut (PropertyGet), source lines 411-413
' -----------------------------------------------------------------------------
Public Property Get OutPut() As String
    OutPut = lOutPut
End Property

' -----------------------------------------------------------------------------
' OutPut (PropertyLet), source lines 414-417
' -----------------------------------------------------------------------------
Public Property Let OutPut(oValue As String)
    If oValue <> lOutPut Then NeedSave = True
    lOutPut = oValue
End Property

' -----------------------------------------------------------------------------
' Device (PropertyGet), source lines 418-420
' -----------------------------------------------------------------------------
Public Property Get Device() As String
    Device = lDevice
End Property

' -----------------------------------------------------------------------------
' Device (PropertyLet), source lines 421-424
' -----------------------------------------------------------------------------
Public Property Let Device(oValue As String)
    If oValue <> lDevice Then NeedSave = True
    lDevice = oValue
End Property

' -----------------------------------------------------------------------------
' file (PropertyGet), source lines 425-427
' -----------------------------------------------------------------------------
Public Property Get file() As String
    file = lFile
End Property

' -----------------------------------------------------------------------------
' AccelerationsCount (PropertyGet), source lines 428-433
' -----------------------------------------------------------------------------
Public Property Get AccelerationsCount() As Long
    If Not Not lAccelerations Then
        AccelerationsCount = UBound(lAccelerations) + 1
        
    End If
End Property

' -----------------------------------------------------------------------------
' Accelerations (PropertyGet), source lines 434-436
' -----------------------------------------------------------------------------
Public Property Get Accelerations(q As Long) As Long
    Accelerations = lAccelerations(q)
End Property

' -----------------------------------------------------------------------------
' AccelerationsAsString (PropertyGet), source lines 437-446
' -----------------------------------------------------------------------------
Public Property Get AccelerationsAsString() As String
    If Not Not lAccelerations Then
        Dim q As Long
        For q = 0 To UBound(lAccelerations) - 1
            AccelerationsAsString = AccelerationsAsString & lAccelerations(q) / 10 & ";"
        Next
        AccelerationsAsString = AccelerationsAsString & lAccelerations(q) / 10
        AccelerationsAsString = Replace(AccelerationsAsString, ",", ".")
    End If
End Property

' -----------------------------------------------------------------------------
' AccelerationsAsString (PropertyLet), source lines 447-461
' -----------------------------------------------------------------------------
Public Property Let AccelerationsAsString(oValue As String)
    oValue = Trim(oValue)
    Dim arrStr() As String
    arrStr = Split(oValue, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(StringToDouble(arrStr(q)) * 10)
        Next
    End If
    If Not CompareArrLong(arr, lAccelerations) Then NeedSave = True
    lAccelerations = arr
End Property

' -----------------------------------------------------------------------------
' SpeedsCount (PropertyGet), source lines 462-466
' -----------------------------------------------------------------------------
Public Property Get SpeedsCount() As Long
    If Not Not lSpeeds Then
        SpeedsCount = UBound(lSpeeds) + 1
    End If
End Property

' -----------------------------------------------------------------------------
' Speeds (PropertyGet), source lines 467-469
' -----------------------------------------------------------------------------
Public Property Get Speeds(q As Long) As Long
    Speeds = lSpeeds(q)
End Property

' -----------------------------------------------------------------------------
' SpeedsAsString (PropertyGet), source lines 470-478
' -----------------------------------------------------------------------------
Public Property Get SpeedsAsString() As String
    If Not Not lSpeeds Then
        Dim q As Long
        For q = 0 To UBound(lSpeeds) - 1
            SpeedsAsString = SpeedsAsString & lSpeeds(q) & ";"
        Next
        SpeedsAsString = SpeedsAsString & lSpeeds(q)
    End If
End Property

' -----------------------------------------------------------------------------
' SpeedsAsString (PropertyLet), source lines 479-493
' -----------------------------------------------------------------------------
Public Property Let SpeedsAsString(oValue As String)
    oValue = Trim(oValue)
    Dim arrStr() As String
    arrStr = Split(oValue, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(arrStr(q))
        Next
    End If
    If Not CompareArrLong(arr, lSpeeds) Then NeedSave = True
    lSpeeds = arr
End Property

' -----------------------------------------------------------------------------
' ForcesCount (PropertyGet), source lines 494-498
' -----------------------------------------------------------------------------
Public Property Get ForcesCount() As Long
    If Not Not lForces Then
        ForcesCount = UBound(lForces) + 1
    End If
End Property

' -----------------------------------------------------------------------------
' Forces (PropertyGet), source lines 499-501
' -----------------------------------------------------------------------------
Public Property Get Forces(i As Long) As Long
    Forces = lForces(i)
End Property

' -----------------------------------------------------------------------------
' ForcesAsString (PropertyGet), source lines 502-510
' -----------------------------------------------------------------------------
Public Property Get ForcesAsString() As String
    If Not Not lForces Then
        Dim q As Long
        For q = 0 To UBound(lForces) - 1
            ForcesAsString = ForcesAsString & lForces(q) & ";"
        Next
        ForcesAsString = ForcesAsString & lForces(q)
    End If
End Property

' -----------------------------------------------------------------------------
' ForcesAsString (PropertyLet), source lines 511-525
' -----------------------------------------------------------------------------
Public Property Let ForcesAsString(oValue As String)
    oValue = Trim(oValue)
    Dim arrStr() As String
    arrStr = Split(oValue, ";")
    If UBound(arrStr) > -1 Then
        Dim arr() As Long
        ReDim arr(UBound(arrStr))
        Dim q As Long
        For q = 0 To UBound(arrStr)
            arr(q) = StringToLong(arrStr(q))
        Next
    End If
    If Not CompareArrLong(arr, lForces) Then NeedSave = True
    lForces = arr
End Property

' -----------------------------------------------------------------------------
' Terminator (PropertyGet), source lines 526-528
' -----------------------------------------------------------------------------
Public Property Get Terminator() As String
    Terminator = Trim(lTerminator)
End Property

' -----------------------------------------------------------------------------
' Terminator (PropertyLet), source lines 529-533
' -----------------------------------------------------------------------------
Property Let Terminator(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lTerminator Then NeedSave = True
    lTerminator = Trim(oValue)
End Property

' -----------------------------------------------------------------------------
' EndWorkFile (PropertyGet), source lines 534-536
' -----------------------------------------------------------------------------
Public Property Get EndWorkFile() As String
    EndWorkFile = Trim(lEndWorkFile)
End Property

' -----------------------------------------------------------------------------
' EndWorkFile (PropertyLet), source lines 537-541
' -----------------------------------------------------------------------------
Property Let EndWorkFile(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lEndWorkFile Then NeedSave = True
    lEndWorkFile = Trim(oValue)
End Property

' -----------------------------------------------------------------------------
' BeginWorkFile (PropertyGet), source lines 542-544
' -----------------------------------------------------------------------------
Public Property Get BeginWorkFile() As String
    BeginWorkFile = Trim(lBeginWorkFile)
End Property

' -----------------------------------------------------------------------------
' BeginWorkFile (PropertyLet), source lines 545-549
' -----------------------------------------------------------------------------
Property Let BeginWorkFile(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lBeginWorkFile Then NeedSave = True
    lBeginWorkFile = Trim(oValue)
End Property

' -----------------------------------------------------------------------------
' Initializer (PropertyGet), source lines 550-552
' -----------------------------------------------------------------------------
Public Property Get Initializer() As String
    Initializer = Trim(lInitializer)
End Property

' -----------------------------------------------------------------------------
' Initializer (PropertyLet), source lines 553-557
' -----------------------------------------------------------------------------
Property Let Initializer(oValue As String)
    oValue = Trim(oValue)
    If oValue <> lInitializer Then NeedSave = True
    lInitializer = Trim(oValue)
End Property

' -----------------------------------------------------------------------------
' ActiveForce (PropertyGet), source lines 558-560
' -----------------------------------------------------------------------------
Public Property Get ActiveForce() As Long
    ActiveForce = lActiveForce
End Property

' -----------------------------------------------------------------------------
' ActiveForce (PropertyLet), source lines 561-564
' -----------------------------------------------------------------------------
Property Let ActiveForce(oValue As Long)
    If oValue <> lActiveForce Then NeedSave = True
    lActiveForce = oValue
End Property

' -----------------------------------------------------------------------------
' ActiveAcceleration (PropertyGet), source lines 565-569
' -----------------------------------------------------------------------------
Public Property Get ActiveAcceleration() As Long
    If lActiveAcceleration > AccelerationsCount Then lActiveAcceleration = AccelerationsCount
    If lActiveAcceleration < 0 Then lActiveAcceleration = 0
    ActiveAcceleration = lActiveAcceleration
End Property

' -----------------------------------------------------------------------------
' ActiveAcceleration (PropertyLet), source lines 570-573
' -----------------------------------------------------------------------------
Property Let ActiveAcceleration(oValue As Long)
    If oValue <> lActiveAcceleration Then NeedSave = True
    lActiveAcceleration = oValue
End Property

' -----------------------------------------------------------------------------
' ActiveSpeed (PropertyGet), source lines 574-578
' -----------------------------------------------------------------------------
Public Property Get ActiveSpeed() As Long
    If lActiveSpeed > SpeedsCount Then lActiveSpeed = SpeedsCount
    If lActiveSpeed < 0 Then lActiveSpeed = 0
    ActiveSpeed = lActiveSpeed
End Property

' -----------------------------------------------------------------------------
' ActiveSpeed (PropertyLet), source lines 579-582
' -----------------------------------------------------------------------------
Property Let ActiveSpeed(oValue As Long)
    If oValue <> lActiveSpeed Then NeedSave = True
    lActiveSpeed = oValue
End Property

' -----------------------------------------------------------------------------
' ActiveSpeedUP (PropertyGet), source lines 583-587
' -----------------------------------------------------------------------------
Public Property Get ActiveSpeedUP() As Long
    If lActiveSpeedUP > SpeedsCount Then lActiveSpeedUP = SpeedsCount
    If lActiveSpeedUP < 0 Then lActiveSpeedUP = 0
    ActiveSpeedUP = lActiveSpeedUP
End Property

' -----------------------------------------------------------------------------
' ActiveSpeedUP (PropertyLet), source lines 588-591
' -----------------------------------------------------------------------------
Property Let ActiveSpeedUP(oValue As Long)
    If oValue <> lActiveSpeedUP Then NeedSave = True
    lActiveSpeedUP = oValue
End Property

' -----------------------------------------------------------------------------
' MinimalVersion (PropertyGet), source lines 592-594
' -----------------------------------------------------------------------------
Public Property Get MinimalVersion() As Long
    MinimalVersion = lMinimalVersion
End Property

' -----------------------------------------------------------------------------
' CommandLanguage (PropertyGet), source lines 599-601
' -----------------------------------------------------------------------------
Public Property Get CommandLanguage() As CommandLanguages
    CommandLanguage = lCommandLanguage
End Property

' -----------------------------------------------------------------------------
' CommandLanguage (PropertyLet), source lines 602-605
' -----------------------------------------------------------------------------
Property Let CommandLanguage(oValue As CommandLanguages)
    If oValue <> lCommandLanguage Then NeedSave = True
    lCommandLanguage = oValue
End Property

' -----------------------------------------------------------------------------
' name (PropertyGet), source lines 606-608
' -----------------------------------------------------------------------------
Public Property Get name() As String
    name = Trim(lName)
End Property

' -----------------------------------------------------------------------------
' name (PropertyLet), source lines 609-612
' -----------------------------------------------------------------------------
Property Let name(oValue As String)
    If Trim(oValue) <> Trim(lName) Then NeedSave = True
    lName = Trim(oValue)
End Property

' -----------------------------------------------------------------------------
' MaxSheetWidth (PropertyGet), source lines 613-615
' -----------------------------------------------------------------------------
Public Property Get MaxSheetWidth() As Long
    MaxSheetWidth = lMaxSheetWidth
End Property

' -----------------------------------------------------------------------------
' MaxSheetWidth (PropertyLet), source lines 616-619
' -----------------------------------------------------------------------------
Property Let MaxSheetWidth(oValue As Long)
    If oValue <> lMaxSheetWidth Then NeedSave = True
    lMaxSheetWidth = oValue
End Property

' -----------------------------------------------------------------------------
' MaxSheetHeight (PropertyGet), source lines 620-622
' -----------------------------------------------------------------------------
Public Property Get MaxSheetHeight() As Long
    MaxSheetHeight = lMaxSheetHeight
End Property

' -----------------------------------------------------------------------------
' MaxSheetHeight (PropertyLet), source lines 623-626
' -----------------------------------------------------------------------------
Property Let MaxSheetHeight(oValue As Long)
    If oValue <> lMaxSheetHeight Then NeedSave = True
    lMaxSheetHeight = oValue
End Property

' -----------------------------------------------------------------------------
' Visible (PropertyGet), source lines 627-629
' -----------------------------------------------------------------------------
Public Property Get Visible() As Boolean
    Visible = lVisible
End Property

' -----------------------------------------------------------------------------
' Visible (PropertyLet), source lines 630-633
' -----------------------------------------------------------------------------
Property Let Visible(oValue As Boolean)
    If oValue <> lVisible Then NeedSave = True
    lVisible = oValue
End Property

' -----------------------------------------------------------------------------
' CompareArrLong (Function), source lines 635-648
' -----------------------------------------------------------------------------
Private Function CompareArrLong(ByRef Arr1() As Long, ByRef arr2() As Long) As Boolean
    If CntArr(Arr1) = CntArr(arr2) Then
        CompareArrLong = True
        If CntArr(Arr1) > 0 Then
            Dim q As Long
            For q = 0 To UBound(Arr1)
                If Arr1(q) <> arr2(q) Then
                    CompareArrLong = False
                    Exit For
                End If
            Next
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' CntArr (Function), source lines 649-653
' -----------------------------------------------------------------------------
Private Function CntArr(ByRef arr() As Long) As Long
    If Not Not arr Then
        CntArr = UBound(arr) + 1
    End If
End Function

' -----------------------------------------------------------------------------
' ValPD (Function), source lines 654-656
' -----------------------------------------------------------------------------
Private Function ValPD(ByRef txt As String) As Double
    ValPD = val(Replace$(Replace$(txt, "%", ""), ",", "."))
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 657-662
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    CountProfiles = CountProfiles - 1
    If NeedSave Then
        Save
    End If
End Sub
