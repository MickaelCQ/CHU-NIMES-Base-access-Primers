Option Compare Database

Private Sub Form_Load()
DoCmd.ShowToolbar "Ribbon", acToolbarNo
Me.ShortcutMenu = False
End Sub


Private Sub Form_Timer()
Msg.Caption = Right(Msg.Caption, Len(Msg.Caption) - 1) & Left(Msg.Caption, 1)
End Sub
