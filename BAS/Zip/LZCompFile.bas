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
