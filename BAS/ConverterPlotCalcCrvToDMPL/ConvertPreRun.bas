Public Function ConvertPreRun(Profile As Profile, ByVal dx As Long, ByVal dy As Long, ByVal AllCurvesHeight As Long, ByVal AllCurvesWidth As Long) As String
    Last_X = -2147483647
    Last_Y = -2147483647
    
    Dim x() As Long, y() As Long
    Dim kx As Double, ky As Double
    Dim xp As Long, yp As Long

    ConvertPreRun = ";:HL0,EC1,P1,V5,A"
       
    With Profile
        Select Case .ZeroPosition
            Case ZeroPositions.RightP
                kx = 0.04 * Abs(.HeightScale)
                ky = 0.04 * Abs(.WidthScale)
                
                ReDim x(4), y(4)
                x(0) = 0
                y(0) = 0
                
                x(1) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                y(1) = 0
                
                x(2) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                y(2) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                
                x(3) = 0
                y(3) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                
                x(4) = 0
                y(4) = 0
            Case ZeroPositions.CenterP
                kx = 0.04 * Abs(.WidthScale)
                ky = 0.04 * Abs(.HeightScale)
                
                ReDim x(5), y(5)
                x(0) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(0) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(1) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(1) = (AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(2) = -(AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(2) = (AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(3) = -(AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(3) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(4) = (AllCurvesWidth / 2 + Profile.KnifeOffset)
                y(4) = -(AllCurvesHeight / 2 + Profile.KnifeOffset)
                
                x(5) = 0
                y(5) = 0
            Case ZeroPositions.LeftP
                kx = 0.04 * Abs(.WidthScale)
                ky = 0.04 * Abs(.HeightScale)
                
                ReDim x(4), y(4)
                x(0) = 0
                y(0) = 0
                
                x(1) = 0
                y(1) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                
                x(2) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                y(2) = (AllCurvesHeight + Profile.KnifeOffset * 2)
                
                x(3) = (AllCurvesWidth + Profile.KnifeOffset * 2)
                y(3) = 0
                
                x(4) = 0
                y(4) = 0
        End Select
    End With
    
    Dim q As Long
    For q = 0 To UBound(x) - 1
        xp = x(q) * kx + dx
        yp = y(q) * ky + dy
        ConvertPreRun = ConvertPreRun & "U" & xp & "," & yp & ","
        If Profile.CornerPoints Then ConvertPreRun = ConvertPreRun & "D"
    Next
    xp = x(q) * kx + dx
    yp = y(q) * ky + dy
    ConvertPreRun = ConvertPreRun & "U" & xp & "," & yp & ","
    ConvertPreRun = ConvertPreRun & "U@"
End Function
