Private Sub UserForm_Terminate()
    BlocksSocket1.Disconnect
    DataFromMeToProfile
    PlotterForm.RefreshFields
End Sub
