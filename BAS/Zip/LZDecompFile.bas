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
