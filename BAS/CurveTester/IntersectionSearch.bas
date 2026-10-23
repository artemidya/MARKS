Public Sub IntersectionSearch()
    On Error GoTo 30
    UserData.Load
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrInch
    
    If Application.ActiveSelection.shapes.count = 0 Then Exit Sub
    
    Application.Optimization = True
    Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("CurveTester.BeginCommandGroup")

    Application.ActiveSelection.ConvertToCurves
    Application.ActiveSelection.UngroupAllEx
    
    
    Dim Curves() As curve
    Dim CountCurves As Long
    
    GetCurves Curves(), CountCurves
    
    ActiveDocument.Undo 1
    
    
    Dim countCrossPoints As Long
    Dim x() As Double, y() As Double
    countCrossPoints = CrossPoints(x(), y(), Curves(), CountCurves)
    

    If countCrossPoints > 0 Then
        If MsgBox(localization.ActiveLng.Item("CurveTester.MsgBox.P1") & " " & countCrossPoints & Chr(10) & Chr(13) & localization.ActiveLng.Item("CurveTester.MsgBox.P2"), vbYesNo, "PlotCalc") = vbYes Then
            DrivePoints x(), y(), countCrossPoints
        End If
    Else
        MsgBox localization.ActiveLng.Item("CurveTester.MsgBox2"), vbInformation, "PlotCalc"
    End If
    Application.ActiveDocument.EndCommandGroup

30
    Application.ActiveDocument.Unit = lastUnit
    Application.Optimization = False
    Application.Refresh
End Sub
