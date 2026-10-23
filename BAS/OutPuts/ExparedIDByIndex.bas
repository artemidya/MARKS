Property Get ExparedIDByIndex(ByVal index As Long) As String
    If (index >= printers.count) And (index < (printers.count + Interfaces.count)) Then
        index = index - printers.count
        ExparedIDByIndex = Interfaces.ExparedIDByIndex(index)
    End If
End Property
