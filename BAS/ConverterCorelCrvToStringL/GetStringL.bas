Public Function GetStringL(shape As VGCore.shape) As String
    'Start Process
    Processing = True
    If Not shape Is Nothing Then
        On Error GoTo 10
        Dim lastUnit As cdrUnit
        lastUnit = Application.ActiveDocument.Unit
        Application.ActiveDocument.Unit = cdrMillimeter
        lastT = Timer
        Percent = 0

        

        GetCurvesCycle shape.shapes, True
        RaiseEvent OnPercentChange(100)
10
        Application.ActiveDocument.Unit = lastUnit
    End If
    GetStringL = OutString
    OutString = ""
    'End Process
    Processing = False
End Function
