' =============================================================================
' Group: 08_Common_Utilities
' Source module: DictionaryC.cls
' Procedures: 5
' =============================================================================

' -----------------------------------------------------------------------------
' Item (PropertyGet), source lines 14-22
' -----------------------------------------------------------------------------
Public Property Get Item(Key As String) As String
    Dim q As Long
    For q = 0 To count - 1
        If Key = Keys(q) Then
            Item = Values(q)
            Exit Function
        End If
    Next
End Property

' -----------------------------------------------------------------------------
' Item (PropertyLet), source lines 23-32
' -----------------------------------------------------------------------------
Public Property Let Item(Key As String, value As String)
    Dim q As Long
    For q = 0 To count - 1
        If Key = Keys(q) Then
            Values(q) = value
            Exit Property
        End If
    Next
    MsgBox "err"
End Property

' -----------------------------------------------------------------------------
' Add (Sub), source lines 34-40
' -----------------------------------------------------------------------------
Public Sub Add(Key As String, value As String)
    ReDim Preserve Keys(count)
    ReDim Preserve Values(count)
    Keys(count) = Key
    Values(count) = value
    count = count + 1
End Sub

' -----------------------------------------------------------------------------
' Key (Function), source lines 42-44
' -----------------------------------------------------------------------------
Public Function Key(index As Long) As String
    Key = Keys(index)
End Function

' -----------------------------------------------------------------------------
' value (Function), source lines 45-47
' -----------------------------------------------------------------------------
Public Function value(index As Long) As String
    value = Values(index)
End Function
