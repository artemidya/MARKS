Private Sub SortCurves()
    Dim i As Long
    Dim fSwitched As Boolean
    
    Dim count As Long
    Dim path As waPath
    count = Curves.count - 1
    
    Dim sum1 As Long
    Dim sum2 As Long
    
    Do
        fSwitched = False
        For i = 1 To count
            
            Set path = Curves(i)
            With path.Bounds
                sum1 = (.right - .left) + (.bottom - .top)
            End With
            
            Set path = Curves(i + 1)
            With path.Bounds
                sum2 = (.right - .left) + (.bottom - .top)
            End With
            
            If sum1 > sum2 Then
                fSwitched = True
                Set path = Curves(i)
                Curves.Remove i
                Curves.Add path, , , i
            End If
        Next
    Loop While fSwitched
End Sub
