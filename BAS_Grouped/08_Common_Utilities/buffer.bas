' =============================================================================
' Group: 08_Common_Utilities
' Source module: buffer.cls
' Procedures: 20
' =============================================================================

' -----------------------------------------------------------------------------
' CRC32 (PropertyGet), source lines 45-53
' -----------------------------------------------------------------------------
Public Property Get CRC32() As Long
    If NeedCalculateCRC Then
        If Not Not my_bytes Then
            lCRC32 = RtlComputeCrc32(0, VarPtr(my_bytes(0)), UBound(my_bytes) + 1)
            NeedCalculateCRC = False
        End If
    End If
    CRC32 = lCRC32
End Property

' -----------------------------------------------------------------------------
' size (PropertyGet), source lines 54-56
' -----------------------------------------------------------------------------
Public Property Get size() As Long
    size = count_bytes
End Property

' -----------------------------------------------------------------------------
' size (PropertyLet), source lines 57-66
' -----------------------------------------------------------------------------
Property Let size(oValue As Long)
    count_bytes = oValue
    If oValue < 1 Then
         Err.Description = "Buffer size = 0"
         Err.Raise 400
    Else
        ReDim my_bytes(oValue - 1)
        NeedCalculateCRC = True
    End If
End Property

' -----------------------------------------------------------------------------
' bytes (PropertyGet), source lines 67-69
' -----------------------------------------------------------------------------
Public Property Get bytes() As Byte()
    bytes = my_bytes
End Property

' -----------------------------------------------------------------------------
' SetBytes (Function), source lines 71-75
' -----------------------------------------------------------------------------
Public Function SetBytes(ByRef bytes() As Byte) As Long
    my_bytes = bytes
    count_bytes = UBound(my_bytes) + 1
    NeedCalculateCRC = True
End Function

' -----------------------------------------------------------------------------
' bPut_Long (Function), source lines 118-128
' -----------------------------------------------------------------------------
Public Function bPut_Long(ByVal value As Long) As Long
    If cp + 4 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        If InterPlatformNetConverter Then value = htonl(value)
        CopyMemory my_bytes(cp), value, 4
        cp = cp + 4
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_Long (Function), source lines 129-138
' -----------------------------------------------------------------------------
Public Function bGet_Long() As Long
    If cp + 4 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Long, my_bytes(cp), 4
        If InterPlatformNetConverter Then bGet_Long = ntohl(bGet_Long)
        cp = cp + 4
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_Integer (Function), source lines 139-149
' -----------------------------------------------------------------------------
Public Function bPut_Integer(ByVal value As Integer) As Long
    If cp + 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        If InterPlatformNetConverter Then value = htons(value)
        CopyMemory my_bytes(cp), value, 2
        cp = cp + 2
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_Integer (Function), source lines 150-159
' -----------------------------------------------------------------------------
Public Function bGet_Integer() As Integer
    If cp + 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Integer, my_bytes(cp), 2
        If InterPlatformNetConverter Then bGet_Integer = ntohs(bGet_Integer)
        cp = cp + 2
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_Boolean (Function), source lines 160-169
' -----------------------------------------------------------------------------
Public Function bPut_Boolean(ByVal value As Boolean) As Long
    If cp + LenB(value) > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory my_bytes(cp), value, LenB(value)
        cp = cp + LenB(value)
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_Boolean (Function), source lines 170-178
' -----------------------------------------------------------------------------
Public Function bGet_Boolean() As Boolean
    If cp + LenB(bGet_Boolean) > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory bGet_Boolean, my_bytes(cp), LenB(bGet_Boolean)
        cp = cp + LenB(bGet_Boolean)
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_String (Function), source lines 179-193
' -----------------------------------------------------------------------------
Public Function bPut_String(ByVal value As String) As Long
    If cp + Len(value) * 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        Dim i As Integer
        For q = 1 To Len(value)
            i = AscW(Mid(value, q, 1))
            CopyMemory my_bytes((q - 1) * 2 + cp), i, 2
        Next
        cp = cp + Len(value) * 2
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_String (Function), source lines 194-207
' -----------------------------------------------------------------------------
Public Function bGet_String(lenSymbols As Long) As String
    If cp + lenSymbols * 2 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        Dim i As Integer
        For q = 1 To lenSymbols
            CopyMemory i, my_bytes((q - 1) * 2 + cp), 2
            bGet_String = bGet_String & ChrW(i)
        Next
        cp = cp + lenSymbols * 2
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_StringASCII (Sub), source lines 208-221
' -----------------------------------------------------------------------------
Public Sub bPut_StringASCII(ByVal value As String)
    If cp + Len(value) > count_bytes Then
        count_bytes = cp + Len(value)
        ReDim Preserve my_bytes(count_bytes - 1)
        bPut_StringASCII value
    Else
        Dim q As Long
        For q = 1 To Len(value)
            my_bytes(cp + q - 1) = Asc(Mid(value, q, 1))
        Next
        cp = cp + Len(value)
        NeedCalculateCRC = True
    End If
End Sub

' -----------------------------------------------------------------------------
' bGet_StringASCII (Function), source lines 222-233
' -----------------------------------------------------------------------------
Public Function bGet_StringASCII(lenSymbols As Long) As String
    If cp + lenSymbols > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim q As Long
        For q = 1 To lenSymbols
            bGet_StringASCII = bGet_StringASCII & Chr(cp + q - 1)
        Next
        cp = cp + lenSymbols
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_Byte (Function), source lines 234-243
' -----------------------------------------------------------------------------
Public Function bPut_Byte(ByVal value As Byte)
    If cp + 1 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        my_bytes(cp) = value
        cp = cp + 1
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_Byte (Function), source lines 244-252
' -----------------------------------------------------------------------------
Public Function bGet_Byte() As Byte
    If cp + 1 > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        bGet_Byte = my_bytes(cp)
        cp = cp + 1
    End If
End Function

' -----------------------------------------------------------------------------
' bPut_Bytes (Function), source lines 253-264
' -----------------------------------------------------------------------------
Public Function bPut_Bytes(ByRef value() As Byte)
    Dim cnt As Long
    cnt = UBound(value) + 1
    If cp + cnt > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        CopyMemory my_bytes(cp), value(0), cnt
        cp = cp + cnt
        NeedCalculateCRC = True
    End If
End Function

' -----------------------------------------------------------------------------
' bGet_Bytes (Function), source lines 265-276
' -----------------------------------------------------------------------------
Public Function bGet_Bytes(cnt As Long) As Byte()
    If cp + cnt > count_bytes Then
        Err.Description = "Out of Buffer (Buffer size=" & count_bytes & ")"
        Err.Raise 404
    Else
        Dim b() As Byte
        ReDim b(cnt - 1)
        CopyMemory b(0), my_bytes(cp), cnt
        bGet_Bytes = b
        cp = cp + cnt
    End If
End Function

' -----------------------------------------------------------------------------
' Class_Terminate (Sub), source lines 278-280
' -----------------------------------------------------------------------------
Private Sub Class_Terminate()
    Erase my_bytes
End Sub
