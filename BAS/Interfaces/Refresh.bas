Public Sub Refresh()
    Class_Terminate 'Deinitialize
    Erase Interfaces()
    lCount = 0
 
    'USB - {a5dcbf10-6530-11d2-901f-00c04fb951ed}
    Dim USB_GUID As guid
    With USB_GUID
        .Data1 = &HA5DCBF10
        .Data2 = &H6530
        .Data3 = &H11D2
        .Data4(0) = &H90
        .Data4(1) = &H1F
        .Data4(2) = &H0
        .Data4(3) = &HC0
        .Data4(4) = &H4F
        .Data4(5) = &HB9
        .Data4(6) = &H51
        .Data4(7) = &HED
    End With
    
    'WireFire {28d78fad-5a12-11d1-ae5b-0000f803a8c2}
    Dim WireFire_GUID As guid
    With WireFire_GUID
        .Data1 = &H28D78FAD
        .Data2 = &H5A12
        .Data3 = &H11D1
        .Data4(0) = &HAE
        .Data4(1) = &H5B
        .Data4(2) = &H0
        .Data4(3) = &H0
        .Data4(4) = &HF8
        .Data4(5) = &H3
        .Data4(6) = &HA8
        .Data4(7) = &HC2
    End With
    
    'Bluetooth
    '{e0cbf06c-cd8b-4647-bb8a-263b43f0f974}
    Dim Bluetooth_GUID As guid
    With Bluetooth_GUID
        .Data1 = &HE0CBF06C
        .Data2 = &HCD8B
        .Data3 = &H4647
        .Data4(0) = &HBB
        .Data4(1) = &H8A
        .Data4(2) = &H26
        .Data4(3) = &H3B
        .Data4(4) = &H43
        .Data4(5) = &HF0
        .Data4(6) = &HF9
        .Data4(7) = &H74
    End With
    
    'COM 86E0D1E0-8089-11D0-9CE4-08003E301F73
    Dim Com_GUID As guid
    With Com_GUID
        .Data1 = &H86E0D1E0
        .Data2 = &H8089
        .Data3 = &H11D0
        .Data4(0) = &H9C
        .Data4(1) = &HE4
        .Data4(2) = &H8
        .Data4(3) = &H0
        .Data4(4) = &H3E
        .Data4(5) = &H30
        .Data4(6) = &H1F
        .Data4(7) = &H73
    End With
  
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("USB: ", USB_GUID, "HIDClass,WPD,Bluetooth,Net,Biometric")
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("Bluetooth: ", Bluetooth_GUID)
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("WireFire: ", WireFire_GUID)
    AddArrayToArray Interfaces, GetDeviceInterfacesListForPlotter("COM: ", Com_GUID)
    

    If Not Not Interfaces() Then
        lCount = UBound(Interfaces) + 1
    Else
        lCount = 0
    End If
End Sub
