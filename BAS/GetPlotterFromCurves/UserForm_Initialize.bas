Private Sub UserForm_Initialize()
    If ActiveSelection.shapes.count < 2 Or ActiveSelection.shapes.count > 2 Then
        MsgBox "Select 2 shapes!", vbInformation: End
    End If
    
    If (ActiveSelection.shapes.All.RightX > 0 And ActiveSelection.shapes.All.LeftX > 0) Or (ActiveSelection.shapes.All.RightX < 0 And ActiveSelection.shapes.All.LeftX < 0) Then
        MsgBox "Set (0,0) Point correctly!", vbInformation: End
    End If
    
    If (ActiveSelection.shapes.All.TopY > 0 And ActiveSelection.shapes.All.BottomY > 0) Or (ActiveSelection.shapes.All.TopY < 0 And ActiveSelection.shapes.All.BottomY < 0) Then
        MsgBox "Set (0,0) Point correctly!", vbInformation: End
    End If
    
    Set Converter = New ConverterCorelCrvToStringL
    Dim plotterStr As String, caretteStr As String
    
    If ActiveSelection.shapes(1).SizeWidth > ActiveSelection.shapes(2).SizeWidth Then
        plotterStr = Converter.GetStringL(ActiveSelection.shapes(1))
        caretteStr = Converter.GetStringL(ActiveSelection.shapes(2))
    Else
        plotterStr = Converter.GetStringL(ActiveSelection.shapes(2))
        caretteStr = Converter.GetStringL(ActiveSelection.shapes(1))
    End If
    
    TextBox1.text = "<PlotterImg>" & plotterStr & "</PlotterImg>" & vbNewLine & vbNewLine & "<CarriageImg>" & caretteStr & "</CarriageImg>"
End Sub
