Private Function ahtCommonFileOpenSave( _
            ByVal hwnd As Long, _
            Optional ByRef Flags As Variant, _
            Optional ByVal InitialDir As Variant, _
            Optional ByVal filter As Variant, _
            Optional ByVal FilterIndex As Variant, _
            Optional ByVal DefaultExt As Variant, _
            Optional ByVal fileName As Variant, _
            Optional ByVal DialogTitle As Variant, _
            Optional ByVal OpenFile As Variant) As Variant

    ' This is the entry point you'll use to call the common
    ' file open/save dialog. The parameters are listed
    ' below, and all are optional.
    '
    ' In:
    ' Flags: one or more of the ahtOFN_* constants, OR'd together.
    ' InitialDir: the directory in which to first look
    ' Filter: a set of file filters, set up by calling
    ' AddFilterItem. See examples.
    ' FilterIndex: 1-based integer indicating which filter
    ' set to use, by default (1 if unspecified)
    ' DefaultExt: Extension to use if the user doesn't enter one.
    ' Only useful on file saves.
    ' FileName: Default value for the file name text box.
    ' DialogTitle: Title for the dialog.
    ' hWnd: parent window handle
    ' OpenFile: Boolean(True=Open File/False=Save As)
    ' Out:
    ' Return Value: Either Null or the selected filename
    Dim OFN As tagOPENFILENAME
    Dim strFileName As String
    Dim strFileTitle As String
    Dim fResult As Boolean

    ' Give the dialog a caption title.
    If IsMissing(InitialDir) Then InitialDir = CurDir
    If IsMissing(filter) Then filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(Flags) Then Flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(fileName) Then fileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(OpenFile) Then OpenFile = True
    ' Allocate string space for the returned strings.
    strFileName = left(fileName & String(256, 0), 256)
    strFileTitle = String(256, 0)
    ' Set up the data structure before you call the function
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = hwnd
        .strFilter = filter
        .nFilterIndex = FilterIndex
        .strFile = strFileName
        .nMaxFile = Len(strFileName)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .Flags = Flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir
        ' Didn't think most people would want to deal with
        ' these options.
        .hInstance = 0
        '.strCustomFilter = ""
        '.nMaxCustFilter = 0
        .lpfnHook = 0
        'New for NT 4.0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
    End With
    ' This will pass the desired data structure to the
    ' Windows API, which will in turn it uses to display
    ' the Open/Save As Dialog.
    If OpenFile Then
        fResult = aht_apiGetOpenFileName(OFN)
    Else
        fResult = aht_apiGetSaveFileName(OFN)
    End If

    ' The function call filled in the strFileTitle member
    ' of the structure. You'll have to write special code
    ' to retrieve that if you're interested.
    If fResult Then
        ' You might care to check the Flags member of the
        ' structure to get information about the chosen file.
        ' In this example, if you bothered to pass in a
        ' value for Flags, we'll fill it in with the outgoing
        ' Flags value.
        If Not IsMissing(Flags) Then Flags = OFN.Flags
        If Flags And ahtOFN_ALLOWMULTISELECT Then
            ' Return the full array.
            Dim items As Variant
            Dim value As String
            value = OFN.strFile
            ' Get rid of empty items:
            Dim i As Integer
            For i = Len(value) To 1 Step -1
              If Mid$(value, i, 1) <> Chr$(0) Then
                Exit For
              End If
            Next i
            value = Mid(value, 1, i)

            ' Break the list up at null characters:
            items = Split(value, Chr(0))

            ' Loop through the items in the "array",
            ' and build full file names:
            Dim numItems As Integer
            Dim result() As String

            numItems = UBound(items) + 1
            If numItems > 1 Then
                ReDim result(0 To numItems - 2)
                For i = 1 To numItems - 1
                    result(i - 1) = FixPath(items(0)) & items(i)
                Next i
                ahtCommonFileOpenSave = result
            Else
                ' If you only select a single item,
                ' Windows just places it in item 0.
                ahtCommonFileOpenSave = items(0)
            End If
        Else
            ahtCommonFileOpenSave = TrimNull(OFN.strFile)
        End If
    Else
        ahtCommonFileOpenSave = vbNullString
    End If
End Function
