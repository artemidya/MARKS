Property Get ExparedID(OutPutName As String) As String
    Dim index As Long
    index = GetIndex(OutPutName)

    With Interfaces(index)
        If Not .isComPortAdapter Then
            ExparedID = .ExparedID
        End If
    End With
End Property
