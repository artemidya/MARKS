Public Property Get Connected(ByVal i As Long) As Boolean
    Connected = printers(i).Active
End Property
