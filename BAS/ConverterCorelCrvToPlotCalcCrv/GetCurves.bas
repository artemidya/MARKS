Public Function GetCurves(ByRef Height As Long, ByRef Width As Long, ByRef MaxColor As RGBA, Optional ByRef colors As Collection) As Collection

    'Start Process
    Processing = True
    Height = 0
    Width = 0
    Set Curves = New Collection
    Set MyColors = New Collection
    On Error GoTo 20
    If Not Application.ActiveSelectionRange Is Nothing Then
        On Error GoTo 10
        Dim lastUnit As cdrUnit
        lastUnit = Application.ActiveDocument.Unit
        Application.ActiveDocument.Unit = cdrMillimeter
        
        Height = Application.ActiveSelection.SizeHeight * 1000
        Width = Application.ActiveSelection.SizeWidth * 1000

        lastT = Timer
        Percent = 0
        
        Dim mx As New waMatrix
        mx.Translate -Application.ActiveSelection.LeftX * 1000, -Application.ActiveSelection.BottomY * 1000
        GetCurvesCycle Application.ActiveSelectionRange.shapes, mx, True
        
        RaiseEvent OnPercentChange(100)
        
        
        
        
10
        Application.ActiveDocument.Unit = lastUnit
    End If
20
    Set GetCurves = Curves
    Set colors = MyColors
    MaxColor = MaxColorF
     
    Set Curves = Nothing
    Set MyColors = Nothing
    'End Process
    Processing = False
End Function
