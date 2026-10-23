Public Sub GetCurves(shapes As VGCore.shapes)
    On Error GoTo 10
    Dim lastUnit As cdrUnit
    lastUnit = Application.ActiveDocument.Unit
    Application.ActiveDocument.Unit = cdrMillimeter

       
    TextBox1 = "Private Sub InitCurves()" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Dim myCurves As New Collection" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Set xxxCol = myCurves" & vbNewLine
    TextBox1 = TextBox1 & vbTab & "Dim path As waPath" & vbNewLine
    
    GetCurvesCycle shapes

    TextBox1 = TextBox1 & "End Sub" & vbNewLine
10
    Application.ActiveDocument.Unit = lastUnit
End Sub
