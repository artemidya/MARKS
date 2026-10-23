Private Sub FindPostion(ByRef PositionX As Double, ByRef PositionY As Double, SingMaterialWith As Double, SingMaterialHeight As Double, rightGo As Boolean)
    Dim q As Long, chang As Boolean
    Do
        chang = False
        For q = 1 To VGCore.ActiveDocument.ActivePage.shapes.count
            With VGCore.ActiveDocument.ActivePage.shapes(q)
                If .Type <> cdrGuidelineShape Then
                    If (PositionX < .PositionX + .SizeWidth And PositionX + SingMaterialWith > .PositionX) _
                        And (PositionY < .PositionY And PositionY + SingMaterialWith > .PositionY - .SizeHeight) Then
                        chang = True
                        If rightGo = True Then
                            PositionX = .PositionX + .SizeWidth + SingMaterialWith / 20
                        Else
                            PositionY = .PositionY + SingMaterialWith / 20
                        End If
                    End If
                End If
            End With
        Next
    Loop While chang
End Sub
