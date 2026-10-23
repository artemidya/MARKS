Private Function CompareMatrixes(m1 As MatrixForCheck, m2 As VGCore.shape) As Boolean
    With m2.TransformationMatrix
        If m1.d11 <> .d11 Then Exit Function
        If m1.d12 <> .d12 Then Exit Function
        If m1.d21 <> .d21 Then Exit Function
        If m1.d22 <> .d22 Then Exit Function
    End With
    CompareMatrixes = True
End Function
