Private Function Length(ByRef crv As VGCore.curve) As Double
    Dim q As Long
     For q = 1 To crv.Subpaths.count
        Length = Length + crv.Subpaths(q).Length
     Next
End Function
