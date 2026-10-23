' =============================================================================
' Group: 08_Common_Utilities
' Source module: Zip.bas
' Procedures: 15
' =============================================================================

' -----------------------------------------------------------------------------
' Compress (Function), source lines 63-83
' -----------------------------------------------------------------------------
Public Function Compress(str As String) As String
OutString = "": pos = 0
InString = str
  'Dim FileNumber  As Integer
  'Dim FileNumber2 As Integer

    'If LenB(LoadName) = 0 Then
    '    LoadName = m_LoadFilename
    'End If
    'If LenB(SaveName) = 0 Then
    '    SaveName = m_SaveFilename
    'End If
    'FileNumber = FreeFile
    'Open LoadName For Binary As FileNumber
    'FileNumber2 = FreeFile
    'Open SaveName For Binary As FileNumber2
    LZCompFile 'FileNumber, FileNumber2
    'Close FileNumber
    'Close FileNumber2
    Compress = OutString
End Function

' -----------------------------------------------------------------------------
' Expand (Function), source lines 84-106
' -----------------------------------------------------------------------------
Public Function Expand(str As String) As String

OutString = "": pos = 0
InString = str

  'Dim FileNumber  As Integer
  'Dim FileNumber2 As Integer

    'If LenB(LoadName) = 0 Then
    '    LoadName = m_LoadFilename
    'End If
    'If LenB(SaveName) = 0 Then
    '    SaveName = m_SaveFilename
    'End If
   ' FileNumber = FreeFile
   ' Open LoadName For Binary As FileNumber
   ' FileNumber2 = FreeFile
   ' Open SaveName For Binary As FileNumber2
    LZDecompFile 'FileNumber, FileNumber2
    'Close FileNumber
    'Close FileNumber2
Expand = OutString
End Function

' -----------------------------------------------------------------------------
' DeGetByte (Function), source lines 108-124
' -----------------------------------------------------------------------------
Private Function DeGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = bufferSize Then
        ReDim InBuffer(bufferSize - 1)
        'Get Infile, , InBuffer
        StrToBite InBuffer
        
        BufferPointer& = 0
        DeGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        DeGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function

' -----------------------------------------------------------------------------
' DePutByte (Sub), source lines 126-143
' -----------------------------------------------------------------------------
Private Sub DePutByte(ByVal TheByte As Byte)

  Dim WriteBuffer() As Byte

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = bufferSize - 1 Then
        OutBuffer(bufferSize - 1) = TheByte
        ReDim WriteBuffer(OutPointer)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer + 1
        'Put Outfile, , WriteBuffer
        OutString = OutString & BiteToStr(WriteBuffer())
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub

' -----------------------------------------------------------------------------
' EnGetByte (Function), source lines 145-158
' -----------------------------------------------------------------------------
Private Function EnGetByte() As Byte

    InFilePointer& = InFilePointer& + 1
    If InPointer& = bufferSize Then
        'Get Infile, , InBuffer
        StrToBite InBuffer
        EnGetByte = InBuffer(0)
        InPointer& = 1
     Else 'NOT INPOINTER&...
        EnGetByte = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
    End If

End Function

' -----------------------------------------------------------------------------
' EnPutByte (Sub), source lines 160-174
' -----------------------------------------------------------------------------
Private Sub EnPutByte(ByVal TheByte As Byte)

    OutFilePointer& = OutFilePointer& + 1
    If OutPointer& = bufferSize - 1 Then
        OutBuffer(bufferSize - 1) = TheByte
        BufferPointer& = 0
        'Put Outfile, , OutBuffer
         OutString = OutString & BiteToStr(OutBuffer())
        OutPointer& = 0
     Else 'NOT OUTPOINTER&...
        OutBuffer(OutPointer&) = TheByte
        OutPointer& = OutPointer& + 1
    End If

End Sub

' -----------------------------------------------------------------------------
' FDeleteNode (Sub), source lines 180-213
' -----------------------------------------------------------------------------
Private Sub FDeleteNode(ByVal p As Integer)

  Dim q As Integer

    If dad(p) = NIL Then
        Exit Sub
    End If
    If rson(p) = NIL Then
        q = lson(p)
     ElseIf lson(p) = NIL Then 'NOT RSON(P)...
        q = rson(p)
     Else 'NOT LSON(P)...
        q = lson(p)
        If rson(q) <> NIL Then
            Do
                q = rson(q)
            Loop While rson(q) <> NIL
            rson(dad(q)) = lson(q)
            dad(lson(q)) = dad(q)
            lson(q) = lson(p)
            dad(lson(p)) = q
        End If
        rson(q) = rson(p)
        dad(rson(p)) = q
    End If
    dad(q) = dad(p)
    If rson(dad(p)) = p Then
        rson(dad(p)) = q
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = q
    End If
    dad(p) = NIL

End Sub

' -----------------------------------------------------------------------------
' FInitTree (Sub), source lines 215-226
' -----------------------------------------------------------------------------
Private Sub FInitTree()

  Dim i As Integer

    For i = n + 1 To (n + 256)
        rson(i) = NIL
    Next i
    For i = 0 To n - 1
        dad(i) = NIL
    Next i

End Sub

' -----------------------------------------------------------------------------
' FInsertNode (Sub), source lines 228-289
' -----------------------------------------------------------------------------
Private Sub FInsertNode(ByVal R As Integer)
  Dim i   As Integer
  Dim p   As Integer
  Dim cmp As Integer
  Dim Key As Integer
  Dim x   As Long
  Dim x1  As Integer
  Dim x2  As Integer
    cmp = 1
    Key = R
    p = n + 1 + textbuf(R)
    rson(R) = NIL
    lson(R) = NIL
    matchlength = 0
    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If cmp >= 0 Then
            If rson(p) <> NIL Then
                p = rson(p)
             Else 'NOT RSON(P)...
                rson(p) = R
                dad(R) = p
                Exit Sub
            End If
         Else 'NOT CMP...
            If lson(p) <> NIL Then
                p = lson(p)
             Else 'NOT LSON(P)...
                lson(p) = R
                dad(R) = p
                Exit Sub
            End If
        End If
        For i = 1 To f - 1
            x1 = textbuf(R + i)
            x2 = textbuf(p + i)
            cmp = x1 - x2
            If cmp <> 0 Then
                Exit For
            End If
        Next i
        If i > matchlength Then
            matchposition = p
            matchlength = i
            If matchlength >= f Then
                Exit For
            End If
        End If
    Next x
    dad(R) = dad(p)
    lson(R) = lson(p)
    rson(R) = rson(p)
    dad(lson(p)) = R
    dad(rson(p)) = R
    If rson(dad(p)) = p Then
        rson(dad(p)) = R
     Else 'NOT RSON(DAD(P))...
        lson(dad(p)) = R
    End If
    dad(p) = NIL

End Sub

' -----------------------------------------------------------------------------
' FlushBuffer (Sub), source lines 291-303
' -----------------------------------------------------------------------------
Private Sub FlushBuffer()

  Dim WriteBuffer() As Byte

    If OutPointer& > 0 Then
        ReDim WriteBuffer(OutPointer - 1)
        CopyMemory WriteBuffer(0), OutBuffer(0), OutPointer
        'Put Outfile, , WriteBuffer
        OutString = OutString & BiteToStr(WriteBuffer())
     Else 'NOT OUTPOINTER&...
    End If

End Sub

' -----------------------------------------------------------------------------
' LZCompFile (Sub), source lines 308-418
' -----------------------------------------------------------------------------
Private Sub LZCompFile()

  Dim codebuf(17)     As Byte
  Dim codebufptr      As Integer
  Dim i               As Integer
  Dim cc              As Byte
  Dim R               As Integer
  Dim s               As Integer
  Dim lenn            As Integer
  Dim lastmatchlength As Integer
  Dim Mask            As Byte

    'Infile = file1
    'Outfile = file2
    'DataLen = LOF(Infile)
    DataLen = Len(InString)
    'With Compressheader
    '    .Filesize = DataLen
    '    .Magic = &H44445A53
    '    .Magic2 = &H3327F088
    '    .Magic3 = &H41
    'End With 'Compressheader
    'Put Outfile, , Compressheader
    
    FInitTree
    codebuf(0) = 0
    codebufptr = 1
    Mask = 1
    s = 0
    R = n - f
    For i = s To R - 1
        textbuf(i) = 32
    Next i
    OpenBuffer
    For lenn = 0 To f - 1
        If InFilePointer& = InLen& Then
            Exit For
        End If
        cc = EnGetByte
        textbuf(R + lenn) = cc
    Next lenn
    textsize = lenn
    If textsize = 0 Then
        Exit Sub
    End If
    For i = 1 To f
        FInsertNode (R - i)
    Next i
    FInsertNode (R)
    Do
        If matchlength > lenn Then
            matchlength = lenn
        End If
        If matchlength <= THRESHOLD Then
            matchlength = 1
            codebuf(0) = (codebuf(0) Or Mask)
            codebuf(codebufptr) = textbuf(R)
            codebufptr = codebufptr + 1
         Else 'NOT MATCHLENGTH...
            matchposition = matchposition + 2
            codebuf(codebufptr) = ((matchposition) And 255) 'Achtung normal ohne + 2
            codebufptr = codebufptr + 1
            codebuf(codebufptr) = (((matchposition \ 16) And 240) Or (matchlength - (THRESHOLD + 1))) And 255
            codebufptr = codebufptr + 1
        End If
        Mask = (Mask * 2) And 255
        If Mask = 0 Then
            For i = 0 To codebufptr - 1
                EnPutByte codebuf(i)
            Next i
            codesize = codesize + codebufptr
            codebuf(0) = 0
            Mask = 1
            codebufptr = Mask
        End If
        lastmatchlength = matchlength
        For i = 0 To lastmatchlength - 1
            If InFilePointer& = InLen& Then
                Exit For
            End If
            cc = EnGetByte
            FDeleteNode (s)
            textbuf(s) = cc
            If s < (f - 1) Then
                textbuf(s + n) = cc
            End If
            s = (s + 1) And (n - 1)
            R = (R + 1) And (n - 1)
            FInsertNode (R)
        Next i
        textsize = textsize + i
        Do While i < lastmatchlength
            i = i + 1
            FDeleteNode (s)
            s = (s + 1) And (n - 1)
            R = (R + 1) And (n - 1)
            lenn = lenn - 1
            If lenn <> 0 Then
                FInsertNode (R)
            End If
        Loop
    Loop While lenn > 0
    If codebufptr > 1 Then
        For i = 0 To codebufptr - 1
            EnPutByte codebuf(i)
        Next i
        codesize = codesize + codebufptr
    End If
    FlushBuffer

End Sub

' -----------------------------------------------------------------------------
' LZDecompFile (Sub), source lines 420-493
' -----------------------------------------------------------------------------
Private Sub LZDecompFile()

  Dim Flags     As Byte
  Dim flagcount As Byte
  Dim x         As Long
  Dim bb        As Byte
  Dim b1        As Byte
  Dim b2        As Byte
  Dim i         As Integer
  Dim j         As Integer
  Dim k         As Integer
  Dim R         As Integer

   ' Infile = file1
   ' Outfile = file2
   ' DataLen = LOF(Infile) - 14
    DataLen = Len(InString) ' - 14
    'Get Infile, , Compressheader
    For i = 0 To (n - f) - 1
        textbuf(i) = 32
    Next i
    R = n - f
    Flags = 0
    flagcount = 8
    OpenBuffer
    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If flagcount = 8 Then
            If InFilePointer& = InLen& Then
                Exit For
            End If
            bb = DeGetByte()
            flagcount = 0
            Flags = bb
        End If
        If (Flags And 1) = 1 Then
            If InFilePointer& = InLen& Then
                Exit For
            End If
            bb = DeGetByte()
            DePutByte bb
            textbuf(R) = bb
            R = R + 1
            R = (R And (n - 1))
         Else 'NOT (FLAGS...
            If InFilePointer& = InLen& Then
                Exit For
            End If
            b1 = DeGetByte()
            b1 = b1
            b! = b1
            If InFilePointer& = InLen& Then
                Exit For
            End If
            b2 = DeGetByte()
            i = b1
            j = b2
            i = i Or ((j And 240) * 16) 'Standort der Übereinstimmung
            i = i - 2 'Achtung nomal nicht - 2
            j = (j And 15) + THRESHOLD 'Länge der Übereinstimmung
            For k = 0 To j
                bb = textbuf((i + k) And (n - 1))
                DePutByte bb
                textbuf(R) = bb
                R = R + 1
                R = (R And (n - 1))
            Next k
        End If
        Flags = (Flags \ 2)
        flagcount = flagcount + 1
    Next x
    FlushBuffer

End Sub

' -----------------------------------------------------------------------------
' OpenBuffer (Sub), source lines 495-507
' -----------------------------------------------------------------------------
Private Sub OpenBuffer()

    ReDim InBuffer(bufferSize - 1) As Byte
    ReDim OutBuffer(bufferSize - 1) As Byte
    InLen& = DataLen&
   ' Get Infile, , InBuffer
    StrToBite InBuffer
    InPointer& = 0
    InFilePointer& = 0
    OutPointer& = 0
    OutFilePointer& = 0

End Sub

' -----------------------------------------------------------------------------
' BiteToStr (Function), source lines 511-516
' -----------------------------------------------------------------------------
Private Function BiteToStr(ByRef m() As Byte) As String
Dim q As Long
For q = 0 To UBound(m)
    BiteToStr = BiteToStr & Chr(m(q))
Next
End Function

' -----------------------------------------------------------------------------
' StrToBite (Function), source lines 517-525
' -----------------------------------------------------------------------------
Private Function StrToBite(ByRef b() As Byte) As Byte
On Error GoTo 10
Dim q As Long
    For q = 0 To UBound(b)
        b(q) = Asc(Mid(InString, pos + 1, 1))
        pos = pos + 1
    Next
10
End Function
