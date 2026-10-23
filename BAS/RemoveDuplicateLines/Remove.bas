Public Sub Remove()
    If Processing Then Exit Sub
    UserData.Load
    LastRefreshTime = Timer
    StartTime = LastRefreshTime
    Processing = True
    ProgressFormIsShowed = False
    ThresholdAngle = 1.74532925199433E-02 * 5
    
    lastUnit = Application.ActiveDocument.Unit
    On Error GoTo er:
    Application.ActiveDocument.Unit = cdrMillimeter
    Dim RegData1 As New RegData

    Sensitivity = CDbl(RegData1.Sensitivity) / 1000
    OutlineUse = RegData1.OutlineUse

    Aborted = False
    Application.Status.BeginProgress
    Application.Optimization = True
    
    Application.ActiveDocument.PreserveSelection = True

    Dim srToRemoveDuplicates As VGCore.ShapeRange
    Set srToRemoveDuplicates = GetShapesRage(VGCore.ActiveSelectionRange)
        
    If srToRemoveDuplicates.count > 0 Then
        On Error GoTo e:
        Application.ActiveDocument.BeginCommandGroup localization.ActiveLng.Item("RemoveDuplicateLines.Command")
        Dim Curves() As VGCore.curve
        currentOperationNumber = 0
        OperationCount = 0

        Curves = GetCurves(srToRemoveDuplicates)
        Dim sr As New ShapeRange
        sr.AddRange VGCore.ActiveSelectionRange
        Application.ActiveDocument.ClearSelection
        
        
        t = Timer
'        Dim t2 As Double
'        t2 = Timer
        RemoveCoincidentCurvesSelfy Curves()
        
        ProgressForm.Percent = 1
        Application.Status.Progress = 50
        currentOperationNumber = 0
        OperationCount = (UBound(Curves()) + 1) * UBound(Curves()) \ 2
        RemoveCoincidentCurves Curves()
        
        ProgressForm.Percent = 1
        Application.Status.Progress = 100
        
        DeleteZerro Curves()
        Connecting Curves()
        closing Curves()
        sr.All.AddToSelection
        
e:
        ProgressForm.MbClosing = True
        Unload ProgressForm
        Application.ActiveDocument.EndCommandGroup
    End If

er:
    On Error Resume Next
    Application.ActiveDocument.Unit = lastUnit
    Application.ActiveDocument.PreserveSelection = False
    
    If Aborted Then Application.ActiveDocument.Undo
    Application.Optimization = False
    Application.Status.EndProgress
    Processing = False
    Application.Refresh

'    MsgBox Timer - t2

End Sub
