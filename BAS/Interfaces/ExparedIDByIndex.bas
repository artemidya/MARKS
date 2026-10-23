Property Get ExparedIDByIndex(ByVal index As Long) As String
    With Interfaces(index)
        If Not .isComPortAdapter Then
            ExparedIDByIndex = .ExparedID
        End If
    End With
End Property
