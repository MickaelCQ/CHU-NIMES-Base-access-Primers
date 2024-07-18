Option Compare Database

'Afficher ruban, activer clic-droit, activer le mode création, activer le volet de navigation :
Private Sub Bouton_administrateur_Click()
Dim Msg As String
Msg = " vous êtes en mode administrateur, attention certaines manipulations peuvent modifier la structure de la plateforme !"
DoCmd.ShowToolbar "Ribbon", acToolbarYes
Me.ShortcutMenu = True
Application.CommandBars("View").Controls("Mode créati&on").Enabled = True
End Sub

'Masquer ruban, désactiver clic-droit, désactiver mode création, verouiller volet de navigation :

Private Sub Bouton_utilisateur_Click()
DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide
DoCmd.ShowToolbar "Ribbon", acToolbarNo
Me.ShortcutMenu = False
Application.CommandBars("View").Controls("Mode créati&on").Enabled = False
End Sub

Private Sub Consulter_CQI_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Administrateur : Fonctionnalité en cours de développement "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information fonctionnalité CQI"    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Btn2_Ouvrir_eurofins_Click()
Dim strString As String
strString = "https://eurofinsgenomics.eu/en/ecom/oligonucleotides/custom-dna-oligos/custom-dna-oligos/"
Application.FollowHyperlink strString
End Sub

Private Sub Btn2_Ouvrir_Primer3Plus_Click()
Dim strString As String
strString = "https://www.bioinformatics.nl/cgi-bin/primer3plus/primer3plus.cgi"
Application.FollowHyperlink strString
End Sub

Private Sub Btn2_Ouvrir_SnpCheck_Click_Click()
Dim strString As String
strString = "https://genetools.org/SNPCheck/snpcheck.htm"
Application.FollowHyperlink strString
End Sub

'Masquage par défaut des menus contextuels et du ruban de travail au démarrage :

Private Sub Form_Load()
DoCmd.ShowToolbar "Ribbon", acToolbarNo
Me.ShortcutMenu = False
End Sub

Private Sub Form_Timer()
Zone_en_developpement.Caption = Right(Zone_en_developpement.Caption, Len(Zone_en_developpement.Caption) - 1) & Left(Zone_en_developpement.Caption, 1)
End Sub



