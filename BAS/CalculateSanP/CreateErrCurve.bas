Private Function CreateErrCurve(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As curve
    Set CreateErrCurve = New curve
    Dim ce(1 To 9) As CurveElement
    ce(1).ElementType = cdrElementStart
    ce(1).PositionX = x1
    ce(1).PositionY = y1
    ce(2).ElementType = cdrElementLine
    ce(2).PositionX = x2
    ce(2).PositionY = y2
    
    ce(3).ElementType = cdrElementStart
    ce(3).PositionX = x1
    ce(3).PositionY = y2
    ce(4).ElementType = cdrElementLine
    ce(4).PositionX = x2
    ce(4).PositionY = y1
  '===================================
    ce(5).ElementType = cdrElementStart
    ce(5).PositionX = x1
    ce(5).PositionY = y2
    
    ce(6).ElementType = cdrElementLine
    ce(6).PositionX = x2
    ce(6).PositionY = y2
    
    ce(7).ElementType = cdrElementLine
    ce(7).PositionX = x2
    ce(7).PositionY = y1
    
    ce(8).ElementType = cdrElementLine
    ce(8).PositionX = x1
    ce(8).PositionY = y1
    
    ce(9).ElementType = cdrElementLine
    ce(9).PositionX = x1
    ce(9).PositionY = y2
    
    CreateErrCurve.CreateSubPathFromArray ce
End Function
