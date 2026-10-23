Private Sub RefreshXY()
    Label18.Caption = "X = " & CStr(Round(OutPuts.Xmm, 0)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    Label19.Caption = "Y = " & CStr(Round(OutPuts.Ymm, 0)) & " " & localization.ActiveLng.Item("RaskroyForm.mm")
    
    If Not PltDocument1(TabStrip1.value) Is Nothing Then
        Dim p As POINTL
        p.x = OutPuts.Xmm * 1000
        p.y = OutPuts.Ymm * 1000
        PltDocument1(TabStrip1.value).XYmargin = p
    End If

'    If OutPuts.x = 0 And OutPuts.y = 0 Then
'        CommandButton21.Enabled = False
'    Else
'        CommandButton21.Enabled = True
'    End If
End Sub
