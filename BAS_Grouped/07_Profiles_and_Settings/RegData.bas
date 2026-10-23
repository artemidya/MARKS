' =============================================================================
' Group: 07_Profiles_and_Settings
' Source module: RegData.cls
' Procedures: 38
' =============================================================================

' -----------------------------------------------------------------------------
' NestingPlaceOnPage (PropertyGet), source lines 12-14
' -----------------------------------------------------------------------------
Property Get NestingPlaceOnPage() As Boolean
    If GetSetting(AppName, "Nesting", "NestingPlaceOnPage") = "1" Then NestingPlaceOnPage = True
End Property

' -----------------------------------------------------------------------------
' IP (PropertyLet), source lines 15-19
' -----------------------------------------------------------------------------
Property Let IP(ProfileNumber As String, oValue As String)
    On Error GoTo 10
    SaveSetting AppName, "Profile" & ProfileNumber, "IP", Trim(oValue)
10
End Property

' -----------------------------------------------------------------------------
' IP (PropertyGet), source lines 20-25
' -----------------------------------------------------------------------------
Property Get IP(ProfileNumber As String) As String
    IP = Trim(GetSetting(AppName, "Profile" & ProfileNumber, "IP"))
    If IP = "" Then
        IP = "192.168.1.1"
    End If
End Property

' -----------------------------------------------------------------------------
' Port (PropertyLet), source lines 26-30
' -----------------------------------------------------------------------------
Property Let Port(ProfileNumber As String, oValue As Integer)
    On Error GoTo 10
    SaveSetting AppName, "Profile" & ProfileNumber, "Port", oValue
10
End Property

' -----------------------------------------------------------------------------
' Port (PropertyGet), source lines 31-40
' -----------------------------------------------------------------------------
Property Get Port(ProfileNumber As String) As Integer
    On Error GoTo 10
    Port = CInt(Trim(GetSetting(AppName, "Profile" & ProfileNumber, "Port")))
    

10
    If Port = 0 Then
        Port = 9100
    End If
End Property

' -----------------------------------------------------------------------------
' NestingPlaceOnPage (PropertyLet), source lines 42-48
' -----------------------------------------------------------------------------
Property Let NestingPlaceOnPage(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Nesting", "NestingPlaceOnPage", "1"
    Else
        SaveSetting AppName, "Nesting", "NestingPlaceOnPage", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' ComSettingsStr (PropertyLet), source lines 50-54
' -----------------------------------------------------------------------------
Property Let ComSettingsStr(ComPortNumber As String, oValue As String)
    On Error GoTo 10
    SaveSetting AppName, "COM", UCase(Trim(ComPortNumber)), oValue
10
End Property

' -----------------------------------------------------------------------------
' ComSettingsStr (PropertyGet), source lines 55-64
' -----------------------------------------------------------------------------
Property Get ComSettingsStr(ComPortNumber As String) As String
    On Error GoTo 10
     ComSettingsStr = Trim(GetSetting(AppName, "COM", UCase(Trim(ComPortNumber))))
     If ComSettingsStr = "" Then
        ComSettingsStr = "baud=9600 parity=N data=8 stop=1 xon=off odsr=off dtr=on rts=on"
     End If
     Exit Property
10
    ComSettingsStr = "baud=9600 parity=N data=8 stop=1 xon=off odsr=off dtr=on rts=on"
End Property

' -----------------------------------------------------------------------------
' FirstCutDateLng (PropertyGet), source lines 67-71
' -----------------------------------------------------------------------------
Property Get FirstCutDateLng() As Long
    On Error GoTo 10
    FirstCutDateLng = CLng(GetSetting(AppName, Section, "FirstCutDateLng"))
10
End Property

' -----------------------------------------------------------------------------
' FirstCutDateLng (PropertyLet), source lines 73-75
' -----------------------------------------------------------------------------
Property Let FirstCutDateLng(oValue As Long)
    SaveSetting AppName, Section, "FirstCutDateLng", oValue
End Property

' -----------------------------------------------------------------------------
' ConnectingCurves (PropertyGet), source lines 76-78
' -----------------------------------------------------------------------------
Property Get ConnectingCurves() As Boolean
    If GetSetting(AppName, "Regroup", "ConnectingCurves") = "1" Then ConnectingCurves = True
End Property

' -----------------------------------------------------------------------------
' ConnectingCurves (PropertyLet), source lines 80-86
' -----------------------------------------------------------------------------
Property Let ConnectingCurves(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Regroup", "ConnectingCurves", "1"
    Else
        SaveSetting AppName, "Regroup", "ConnectingCurves", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' ReduceCurves (PropertyGet), source lines 89-91
' -----------------------------------------------------------------------------
Property Get ReduceCurves() As Boolean
    If GetSetting(AppName, "Regroup", "ReduceCurves") = "1" Then ReduceCurves = True
End Property

' -----------------------------------------------------------------------------
' ReduceCurves (PropertyLet), source lines 93-99
' -----------------------------------------------------------------------------
Property Let ReduceCurves(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Regroup", "ReduceCurves", "1"
    Else
        SaveSetting AppName, "Regroup", "ReduceCurves", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' Sensitivity (PropertyGet), source lines 100-107
' -----------------------------------------------------------------------------
Property Get Sensitivity() As Long
    On Error GoTo 10
    Sensitivity = CLng(GetSetting(AppName, "Duplicates", "Sensitivity"))
    
10:
    If Sensitivity < 25 Then Sensitivity = 25
    If Sensitivity > 10000 Then Sensitivity = 10000
End Property

' -----------------------------------------------------------------------------
' Sensitivity (PropertyLet), source lines 109-113
' -----------------------------------------------------------------------------
Property Let Sensitivity(oValue As Long)
    If oValue < 25 Then oValue = 25
    If oValue > 10000 Then oValue = 10000
    SaveSetting AppName, "Duplicates", "Sensitivity", oValue
End Property

' -----------------------------------------------------------------------------
' OutlineUse (PropertyGet), source lines 114-116
' -----------------------------------------------------------------------------
Property Get OutlineUse() As Boolean
    If GetSetting(AppName, "Duplicates", "OutlineUse") = "1" Then OutlineUse = True
End Property

' -----------------------------------------------------------------------------
' OutlineUse (PropertyLet), source lines 118-124
' -----------------------------------------------------------------------------
Property Let OutlineUse(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "Duplicates", "OutlineUse", "1"
    Else
        SaveSetting AppName, "Duplicates", "OutlineUse", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' DelFill (PropertyGet), source lines 125-127
' -----------------------------------------------------------------------------
Property Get DelFill() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "DelFill") = "1" Then DelFill = True
End Property

' -----------------------------------------------------------------------------
' DelFill (PropertyLet), source lines 128-134
' -----------------------------------------------------------------------------
Property Let DelFill(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "DelFill", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "DelFill", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' DelOutline (PropertyGet), source lines 135-137
' -----------------------------------------------------------------------------
Property Get DelOutline() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "DelOutline") = "1" Then DelOutline = True
End Property

' -----------------------------------------------------------------------------
' DelOutline (PropertyLet), source lines 138-144
' -----------------------------------------------------------------------------
Property Let DelOutline(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "DelOutline", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "DelOutline", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' UngroupAll (PropertyGet), source lines 145-147
' -----------------------------------------------------------------------------
Property Get UngroupAll() As Boolean
    If GetSetting(AppName, "ConvertAllToCurves", "UngroupAll") = "1" Then UngroupAll = True
End Property

' -----------------------------------------------------------------------------
' UngroupAll (PropertyLet), source lines 148-154
' -----------------------------------------------------------------------------
Property Let UngroupAll(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, "ConvertAllToCurves", "UngroupAll", "1"
    Else
        SaveSetting AppName, "ConvertAllToCurves", "UngroupAll", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' CropBMPoffset (PropertyGet), source lines 156-161
' -----------------------------------------------------------------------------
Property Get CropBMPoffset() As Long
    On Error GoTo 10
    CropBMPoffset = CLng(GetSetting(AppName, "CropBMP", "CropBMPoffset"))
    If CropBMPoffset < 0 Then CropBMPoffset = 0
10:
End Property

' -----------------------------------------------------------------------------
' CropBMPoffset (PropertyLet), source lines 163-166
' -----------------------------------------------------------------------------
Property Let CropBMPoffset(oValue As Long)
    If oValue < 0 Then oValue = 0
    SaveSetting AppName, "CropBMP", "CropBMPoffset", oValue
End Property

' -----------------------------------------------------------------------------
' TrimOffset (PropertyGet), source lines 167-172
' -----------------------------------------------------------------------------
Property Get TrimOffset() As Long
    On Error GoTo 10
    TrimOffset = CLng(GetSetting(AppName, "Trim", "TrimOffset"))
    If TrimOffset < 0 Then TrimOffset = 0
10:
End Property

' -----------------------------------------------------------------------------
' TrimOffset (PropertyLet), source lines 174-177
' -----------------------------------------------------------------------------
Property Let TrimOffset(oValue As Long)
    If oValue < 0 Then oValue = 0
    SaveSetting AppName, "Trim", "TrimOffset", oValue
End Property

' -----------------------------------------------------------------------------
' Deployed (PropertyGet), source lines 178-180
' -----------------------------------------------------------------------------
Property Get Deployed() As Boolean
    If GetSetting(AppName, Section, "Deployed") = "1" Then Deployed = True
End Property

' -----------------------------------------------------------------------------
' Deployed (PropertyLet), source lines 182-188
' -----------------------------------------------------------------------------
Property Let Deployed(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, Section, "Deployed", "1"
    Else
        SaveSetting AppName, Section, "Deployed", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' PreRun (PropertyGet), source lines 200-202
' -----------------------------------------------------------------------------
Property Get PreRun() As Boolean
    If GetSetting(AppName, Section, "PreRun") = "1" Then PreRun = True
End Property

' -----------------------------------------------------------------------------
' PreRun (PropertyLet), source lines 203-209
' -----------------------------------------------------------------------------
Property Let PreRun(oValue As Boolean)
    If oValue Then
        SaveSetting AppName, Section, "PreRun", "1"
    Else
        SaveSetting AppName, Section, "PreRun", "0"
    End If
End Property

' -----------------------------------------------------------------------------
' ActiveProfile (PropertyGet), source lines 210-214
' -----------------------------------------------------------------------------
Property Get ActiveProfile() As Long
    On Error GoTo 10
    ActiveProfile = val(GetSetting(AppName, Section, "ActiveProfile"))
10
End Property

' -----------------------------------------------------------------------------
' ActiveProfile (PropertyLet), source lines 215-217
' -----------------------------------------------------------------------------
Property Let ActiveProfile(oValue As Long)
   SaveSetting AppName, Section, "ActiveProfile", CStr(oValue)
End Property

' -----------------------------------------------------------------------------
' LastPlotterDirectory (PropertyGet), source lines 218-222
' -----------------------------------------------------------------------------
Property Get LastPlotterDirectory() As String
    On Error GoTo 10
    LastPlotterDirectory = GetSetting(AppName, Section, "LastPlotterDirectory")
10
End Property

' -----------------------------------------------------------------------------
' LastPlotterDirectory (PropertyLet), source lines 223-225
' -----------------------------------------------------------------------------
Property Let LastPlotterDirectory(oValue As String)
   SaveSetting AppName, Section, "LastPlotterDirectory", oValue
End Property

' -----------------------------------------------------------------------------
' AutoUpdate (PropertyGet), source lines 226-228
' -----------------------------------------------------------------------------
Property Get AutoUpdate() As Boolean
    If GetSetting(AppName, "Settings", "AutoUpdate") = "1" Then AutoUpdate = True
End Property

' -----------------------------------------------------------------------------
' AutoUpdate (PropertyLet), source lines 229-235
' -----------------------------------------------------------------------------
Property Let AutoUpdate(oValue As Boolean)
    If oValue Then
        SaveSetting "PlotCalc", "Settings", "AutoUpdate", "1"
    Else
        SaveSetting "PlotCalc", "Settings", "AutoUpdate", "0"
    End If
End Property
