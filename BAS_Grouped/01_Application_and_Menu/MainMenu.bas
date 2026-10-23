' =============================================================================
' Group: 01_Application_and_Menu
' Source module: MainMenu.bas
' Procedures: 16
' =============================================================================

' -----------------------------------------------------------------------------
' Plot (Sub), source lines 3-11
' -----------------------------------------------------------------------------
Public Sub Plot()
    On Error GoTo 10
'    If COM.PortIsWorking Then
'        MsgBox "The port is busy. Try again later.", vbInformation
'        Exit Sub
'    End If
    PlotterForm.Show 1
10
End Sub

' -----------------------------------------------------------------------------
' Place (Sub), source lines 12-16
' -----------------------------------------------------------------------------
Public Sub Place()
    On Error GoTo 10
    RaskroyForm.Show 1
10
End Sub

' -----------------------------------------------------------------------------
' Circulation (Sub), source lines 17-21
' -----------------------------------------------------------------------------
Public Sub Circulation()
    On Error GoTo 10
    CirculationForm.Show 1
10
End Sub

' -----------------------------------------------------------------------------
' Perimetr (Sub), source lines 22-30
' -----------------------------------------------------------------------------
Public Sub Perimetr()
    On Error GoTo 10
    If CalculateSanP.PerimetrFormVisible Then
        Unload PerimetrForm
    Else
        PerimetrForm.Show 0
    End If
10
End Sub

' -----------------------------------------------------------------------------
' Square (Sub), source lines 31-39
' -----------------------------------------------------------------------------
Public Sub Square()
    On Error GoTo 10
    If CalculateSanP.SquareFormVisible Then
        Unload SquareForm
    Else
        SquareForm.Show 0
    End If
10
End Sub

' -----------------------------------------------------------------------------
' Labirint (Sub), source lines 40-44
' -----------------------------------------------------------------------------
Public Sub Labirint()
    On Error GoTo 10
    LabirintCreator.CreateLabirintShape
10
End Sub

' -----------------------------------------------------------------------------
' ScaleA3 (Sub), source lines 45-49
' -----------------------------------------------------------------------------
Public Sub ScaleA3()
    On Error GoTo 10
    Schema.a3
10
End Sub

' -----------------------------------------------------------------------------
' ScaleA4 (Sub), source lines 50-54
' -----------------------------------------------------------------------------
Public Sub ScaleA4()
    On Error GoTo 10
    Schema.a4
10
End Sub

' -----------------------------------------------------------------------------
' ScaleA5 (Sub), source lines 55-59
' -----------------------------------------------------------------------------
Public Sub ScaleA5()
    On Error GoTo 10
    Schema.a5
10
End Sub

' -----------------------------------------------------------------------------
' IntersectionSearch (Sub), source lines 60-64
' -----------------------------------------------------------------------------
Public Sub IntersectionSearch()
    On Error GoTo 10
    CurveTester.IntersectionSearch
10
End Sub

' -----------------------------------------------------------------------------
' CropBMPsub (Sub), source lines 65-69
' -----------------------------------------------------------------------------
Public Sub CropBMPsub()
    On Error GoTo 10
    CropBMP.Crop
10
End Sub

' -----------------------------------------------------------------------------
' TrimCurvesSub (Sub), source lines 70-74
' -----------------------------------------------------------------------------
Public Sub TrimCurvesSub()
    On Error GoTo 10
    TrimCurves.TrimCrv
10
End Sub

' -----------------------------------------------------------------------------
' ReGroupShapes (Sub), source lines 75-79
' -----------------------------------------------------------------------------
Public Sub ReGroupShapes()
    On Error GoTo 10
    ReGroup.StartReGroup
10
End Sub

' -----------------------------------------------------------------------------
' RemoveDuplicate (Sub), source lines 80-84
' -----------------------------------------------------------------------------
Public Sub RemoveDuplicate()
    On Error GoTo 10
    RemoveDuplicateLines.Remove
10
End Sub

' -----------------------------------------------------------------------------
' ConvertToCurves (Sub), source lines 85-89
' -----------------------------------------------------------------------------
Public Sub ConvertToCurves()
    On Error GoTo 10
    ConvertAllToCurves.Convert
10
End Sub

' -----------------------------------------------------------------------------
' settings (Sub), source lines 90-94
' -----------------------------------------------------------------------------
Public Sub settings()
    On Error GoTo 10
    SettingsForm.Show 1
10
End Sub
