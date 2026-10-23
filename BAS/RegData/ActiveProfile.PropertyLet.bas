Property Let ActiveProfile(oValue As Long)
   SaveSetting AppName, Section, "ActiveProfile", CStr(oValue)
End Property
