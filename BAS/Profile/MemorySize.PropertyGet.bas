Public Property Get MemorySize() As Long
    If lMemorySize < 10000 Then MemorySize = 10000
    MemorySize = lMemorySize
End Property
