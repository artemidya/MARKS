' =============================================================================
' Group: 08_Common_Utilities
' Source module: ClipBoard.cls
' Procedures: 1
' =============================================================================

' -----------------------------------------------------------------------------
' SetData (Function), source lines 36-80
' -----------------------------------------------------------------------------
Public Function SetData(MyString As String)
    'PURPOSE: API function to copy text to clipboard
    'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx
    
    #If VBA7 Then
        Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr
        Dim hClipMemory As LongPtr, x As LongPtr
    #Else
        Dim hGlobalMemory As Long, lpGlobalMemory As Long
        Dim hClipMemory As Long, x As Long
    #End If
    
    'Allocate moveable global memory
      hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)
    
    'Lock the block to get a far pointer to this memory.
      lpGlobalMemory = GlobalLock(hGlobalMemory)
    
    'Copy the string to this global memory.
      lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)
    
    'Unlock the memory.
      If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Could not unlock memory location. Copy aborted."
        GoTo OutOfHere2
      End If
    
    'Open the Clipboard to copy data to.
      If OpenClipboard(0&) = 0 Then
        MsgBox "Could not open the Clipboard. Copy aborted."
        Exit Function
      End If
    
    'Clear the Clipboard.
      x = EmptyClipboard()
    
    'Copy the data to the Clipboard.
      hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
    
OutOfHere2:
      If CloseClipboard() = 0 Then
        MsgBox "Could not close Clipboard."
      End If

End Function
