' =============================================================================
' Group: 03_Geometry_and_Curves
' Source module: Schema.bas
' Procedures: 4
' =============================================================================

' -----------------------------------------------------------------------------
' a3 (Sub), source lines 3-87
' -----------------------------------------------------------------------------
Public Sub a3()
    On Error GoTo 10
    UserData.Load
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub
    
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    
    Dim lastReferencePoint As VGCore.cdrReferencePoint
    lastReferencePoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrTopLeft
    
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("Schema.A3.BeginCommandGroup")
    Application.Optimization = True
    
    Dim s1 As ShapeRange
    Dim otstupX As Double, otstupY As Double
    Dim alfa As String, pzoom As String
    alfa = Application.ActiveDocument.fileName
    
    Set s1 = ActiveSelectionRange
    otstupX = s1.PositionX + s1.SizeWidth + 1
    otstupY = s1.PositionY
    
    s1.Duplicate
    s1.ConvertToCurves

    
    NoFill s1

    If s1.SizeWidth > s1.SizeHeight Then s1.Rotate -90#

    Dim page_w As Double, page_h As Double
    page_w = 297
    page_h = 420
    
    If s1.SizeWidth * 25.4 < page_w And s1.SizeHeight * 25.4 < page_h Then
        pzoom = "M 1:1"
    ElseIf s1.SizeWidth * 25.4 / 2 < page_w And s1.SizeHeight * 25.4 / 2 < page_h Then
        pzoom = "M 1:2"
        s1.Stretch 0.5
    ElseIf s1.SizeWidth * 25.4 / 5 < page_w And s1.SizeHeight * 25.4 / 5 < page_h Then
        pzoom = "M 1:5"
        s1.Stretch 0.2
    ElseIf s1.SizeWidth * 25.4 / 10 < page_w And s1.SizeHeight * 25.4 / 10 < page_h Then
        pzoom = "M 1:10"
        s1.Stretch 0.1
    ElseIf s1.SizeWidth * 25.4 / 20 < page_w And s1.SizeHeight * 25.4 / 20 < page_h Then
        pzoom = "M 1:20"
        s1.Stretch 0.05
    ElseIf s1.SizeWidth * 25.4 / 25 < page_w And s1.SizeHeight * 25.4 / 25 < page_h Then
        pzoom = "M 1:25"
        s1.Stretch 0.04
    ElseIf s1.SizeWidth * 25.4 / 50 < page_w And s1.SizeHeight * 25.4 / 50 < page_h Then
        pzoom = "M 1:50"
        s1.Stretch 0.02
    ElseIf s1.SizeWidth * 25.4 / 100 < page_w And s1.SizeHeight * 25.4 / 100 < page_h Then
        pzoom = "M 1:100"
        s1.Stretch 0.001
    Else
        pzoom = "M ?:?"
        Beep
    End If

    Dim x As Double, y As Double
    y = s1.PositionY + 0.6
    x = s1.PositionX
    
    If alfa <> "" Then pzoom = pzoom & " " & alfa
    Application.ActiveLayer.CreateArtisticText x, y, pzoom
    s1.AddToSelection
    
    Application.ActiveSelection.Group
    Application.ActiveSelection.PositionX = otstupX
    Application.ActiveSelection.PositionY = otstupY

    Application.Optimization = False
    Application.ActiveDocument.EndCommandGroup
    Application.Refresh
10
    Application.ActiveDocument.Unit = lastUnit
    ActiveDocument.ReferencePoint = lastReferencePoint
End Sub

' -----------------------------------------------------------------------------
' a4 (Sub), source lines 88-173
' -----------------------------------------------------------------------------
Public Sub a4()
    On Error GoTo 10
    UserData.Load
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub
    
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    Dim lastReferencePoint As VGCore.cdrReferencePoint
    lastReferencePoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrTopLeft
    
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("Schema.A4.BeginCommandGroup")
    Application.Optimization = True
    
    Dim s1 As ShapeRange
    
    Dim otstupX As Double, otstupY As Double
    Dim alfa As String, pzoom As String
    alfa = Application.ActiveDocument.fileName
    
    Set s1 = ActiveSelectionRange
    
    otstupX = s1.PositionX + s1.SizeWidth + 1
    otstupY = s1.PositionY
    
    s1.Duplicate
    s1.ConvertToCurves
    
    NoFill s1

    If s1.SizeWidth > s1.SizeHeight Then s1.Rotate -90#

    Dim page_w As Double, page_h As Double
    page_w = 210
    page_h = 277
    
    If s1.SizeWidth * 25.4 < page_w And s1.SizeHeight * 25.4 < page_h Then
        pzoom = "M 1:1"
    ElseIf s1.SizeWidth * 25.4 / 2 < page_w And s1.SizeHeight * 25.4 / 2 < page_h Then
        pzoom = "M 1:2"
        s1.Stretch 0.5
    ElseIf s1.SizeWidth * 25.4 / 5 < page_w And s1.SizeHeight * 25.4 / 5 < page_h Then
        pzoom = "M 1:5"
        s1.Stretch 0.2
    ElseIf s1.SizeWidth * 25.4 / 10 < page_w And s1.SizeHeight * 25.4 / 10 < page_h Then
        pzoom = "M 1:10"
        s1.Stretch 0.1
    ElseIf s1.SizeWidth * 25.4 / 20 < page_w And s1.SizeHeight * 25.4 / 20 < page_h Then
        pzoom = "M 1:20"
        s1.Stretch 0.05
    ElseIf s1.SizeWidth * 25.4 / 25 < page_w And s1.SizeHeight * 25.4 / 25 < page_h Then
        pzoom = "M 1:25"
        s1.Stretch 0.04
    ElseIf s1.SizeWidth * 25.4 / 50 < page_w And s1.SizeHeight * 25.4 / 50 < page_h Then
        pzoom = "M 1:50"
        s1.Stretch 0.02
    ElseIf s1.SizeWidth * 25.4 / 100 < page_w And s1.SizeHeight * 25.4 / 100 < page_h Then
        pzoom = "M 1:100"
        s1.Stretch 0.001
    Else
        pzoom = "M ?:?"
        Beep
    End If

    Dim x As Double, y As Double
    y = s1.PositionY + 0.6
    x = s1.PositionX
    
    If alfa <> "" Then pzoom = pzoom & " " & alfa
    Application.ActiveLayer.CreateArtisticText x, y, pzoom
    s1.AddToSelection
    
    Application.ActiveSelection.Group
    Application.ActiveSelection.PositionX = otstupX
    Application.ActiveSelection.PositionY = otstupY
    

    Application.Optimization = False
    Application.ActiveDocument.EndCommandGroup
    Application.Refresh
10
    Application.ActiveDocument.Unit = lastUnit
    ActiveDocument.ReferencePoint = lastReferencePoint
End Sub

' -----------------------------------------------------------------------------
' a5 (Sub), source lines 174-256
' -----------------------------------------------------------------------------
Public Sub a5()
    On Error GoTo 10
    UserData.Load
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub
    
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    Dim lastReferencePoint As VGCore.cdrReferencePoint
    lastReferencePoint = ActiveDocument.ReferencePoint
    ActiveDocument.ReferencePoint = cdrTopLeft
    
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("Schema.A5.BeginCommandGroup")
    Application.Optimization = True
    
    Dim s1 As ShapeRange
    Dim otstupX As Double, otstupY As Double
    Dim alfa As String, pzoom As String
    alfa = Application.ActiveDocument.fileName
    
    Set s1 = ActiveSelectionRange
    otstupX = s1.PositionX + s1.SizeWidth + 1
    otstupY = s1.PositionY
    
    s1.Duplicate
    s1.ConvertToCurves
    
    NoFill s1

    If s1.SizeWidth > s1.SizeHeight Then s1.Rotate -90#

    Dim page_w As Double, page_h As Double
    page_w = 148.5
    page_h = 210
    
    If s1.SizeWidth * 25.4 < page_w And s1.SizeHeight * 25.4 < page_h Then
        pzoom = "M 1:1"
    ElseIf s1.SizeWidth * 25.4 / 2 < page_w And s1.SizeHeight * 25.4 / 2 < page_h Then
        pzoom = "M 1:2"
        s1.Stretch 0.5
    ElseIf s1.SizeWidth * 25.4 / 5 < page_w And s1.SizeHeight * 25.4 / 5 < page_h Then
        pzoom = "M 1:5"
        s1.Stretch 0.2
    ElseIf s1.SizeWidth * 25.4 / 10 < page_w And s1.SizeHeight * 25.4 / 10 < page_h Then
        pzoom = "M 1:10"
        s1.Stretch 0.1
    ElseIf s1.SizeWidth * 25.4 / 20 < page_w And s1.SizeHeight * 25.4 / 20 < page_h Then
        pzoom = "M 1:20"
        s1.Stretch 0.05
    ElseIf s1.SizeWidth * 25.4 / 25 < page_w And s1.SizeHeight * 25.4 / 25 < page_h Then
        pzoom = "M 1:25"
        s1.Stretch 0.04
    ElseIf s1.SizeWidth * 25.4 / 50 < page_w And s1.SizeHeight * 25.4 / 50 < page_h Then
        pzoom = "M 1:50"
        s1.Stretch 0.02
    ElseIf s1.SizeWidth * 25.4 / 100 < page_w And s1.SizeHeight * 25.4 / 100 < page_h Then
        pzoom = "M 1:100"
        s1.Stretch 0.001
    Else
        pzoom = "M ?:?"
        Beep
    End If

    Dim x As Double, y As Double
    y = s1.PositionY + 0.6
    x = s1.PositionX
    
    If alfa <> "" Then pzoom = pzoom & " " & alfa
    Application.ActiveLayer.CreateArtisticText x, y, pzoom
    s1.AddToSelection
    
    Application.ActiveSelection.Group
    Application.ActiveSelection.PositionX = otstupX
    Application.ActiveSelection.PositionY = otstupY

    Application.Optimization = False
    Application.ActiveDocument.EndCommandGroup
    Application.Refresh
10
    Application.ActiveDocument.Unit = lastUnit
    ActiveDocument.ReferencePoint = lastReferencePoint
End Sub

' -----------------------------------------------------------------------------
' NoFill (Sub), source lines 257-261
' -----------------------------------------------------------------------------
Private Sub NoFill(ByRef s1 As ShapeRange)
    On Error Resume Next
    s1.ApplyNoFill
    s1.SetOutlineProperties Color:=CreateCMYKColor(0, 0, 0, 100)
End Sub
