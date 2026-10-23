Private Sub Check2_Click()
    If Check2.value = False Then
        Frame3.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        Label7.Enabled = True
        Text7.Enabled = True
        Text8.Enabled = True
        Text12.Enabled = True
        Text13.Enabled = True
        Text14.Enabled = True
        Text15.Enabled = True
        
        Check4.Enabled = True
        Check3.Enabled = True
        Check5.Enabled = True
        Check6.Enabled = True
        Check7.Enabled = True
        Check8.Enabled = True
        Check9.Enabled = True
    Else
        Frame3.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Label5.Enabled = False
        Label6.Enabled = False
        Label7.Enabled = False
        Text7.Enabled = False
        Text8.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
        
        Check4.Enabled = False
        Check3.Enabled = False
        Check5.Enabled = False
        Check6.Enabled = False
        Check7.Enabled = False
        Check8.Enabled = False
        Check9.Enabled = False
    End If
    
    If TextFieldsIsEditing = True Then Exit Sub
    CommandSave.Enabled = True
    NeedSave = True
    Draw ValP(Text4), ValP(Text5), ValP(Text8), ValP(Text7), ValP(Text12), ValP(Text13), ValP(Text14), ValP(Text15)
    RefreshCommand1
End Sub
