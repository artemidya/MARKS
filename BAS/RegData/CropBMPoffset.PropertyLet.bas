Property Let CropBMPoffset(oValue As Long)
    If oValue < 0 Then oValue = 0
    SaveSetting AppName, "CropBMP", "CropBMPoffset", oValue
End Property
