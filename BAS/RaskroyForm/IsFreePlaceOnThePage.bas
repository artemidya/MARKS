Private Function IsFreePlaceOnThePage(ByRef PositionX As Double, ByRef PositionY As Double, SingMaterialWith As Double, SingMaterialHeight As Double) As Boolean
    Dim q As Long

    For q = 1 To VGCore.ActiveDocument.ActivePage.shapes.count
        With VGCore.ActiveDocument.ActivePage.shapes(q)
            If .Type <> cdrGuidelineShape Then
                If (PositionX < .PositionX + .SizeWidth And PositionX + SingMaterialWith > .PositionX) _
                    And (PositionY < .PositionY And PositionY + SingMaterialWith > .PositionY - .SizeHeight) Then
                    If Not FindStaticIdInDu(.StaticId) Then
                        IsFreePlaceOnThePage = False
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    IsFreePlaceOnThePage = True
End Function
