Option Compare Database
Private Sub Btn1_Ouvrir_commandes_Click()
Dim strString As String
strString = "B:\ROSALIND\2-Primers\COMMANDES\Commandes 2022"
Application.FollowHyperlink strString
End Sub

Private Sub Btn1_Ouvrir_eurofins_Click()
Dim strString As String
strString = "https://eurofinsgenomics.eu/en/ecom/oligonucleotides/custom-dna-oligos/custom-dna-oligos/"
Application.FollowHyperlink strString
End Sub

Private Sub Btn1_Ouvrir_Primer3Plus_Click()
Dim strString As String
strString = "https://www.bioinformatics.nl/cgi-bin/primer3plus/primer3plus.cgi"
Application.FollowHyperlink strString
End Sub

Private Sub Btn1_Ouvrir_SnpCheck_Click_Click()
Dim strString As String
strString = "https://genetools.org/SNPCheck/snpcheck.htm"
Application.FollowHyperlink strString
End Sub

Private Sub chkAuteur_Click()

If Me.chkAuteur Then
    Me.txtRechAuteur.Visible = False
Else
    Me.txtRechAuteur.Visible = True
End If

RefreshQuery

End Sub


Private Sub chkFamille_Click()

If Me.chkFamille Then
    Me.cmbRechFamille.Visible = False
Else
    Me.cmbRechFamille.Visible = True
End If

RefreshQuery

End Sub

Private Sub chkResume_Click()

If Me.chkResume Then
    Me.txtRechResume.Visible = False
Else
    Me.txtRechResume.Visible = True
End If

RefreshQuery

End Sub

Private Sub chkTitre_Click()

If Me.chkTitre Then
    Me.txtRechTitre.Visible = False
Else
    Me.txtRechTitre.Visible = True
End If

RefreshQuery

End Sub


Private Sub chkType_Click()

If Me.chkType Then
    Me.cmbRechType.Visible = False
Else
    Me.cmbRechType.Visible = True
End If

RefreshQuery

End Sub

Private Sub cmbRechFamille_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub cmbRechType_AfterUpdate()
Sequence = " Cliquez-ici " & "#B:\ROSALIND\2-Primers\SEQUENCES_DE_REFERENCE\" & [cmbRechType] & ".rtf" & "#"
Me.Requery

End Sub

Private Sub cmbRechType_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub


Private Sub cmbRechType_Click()

End Sub


Private Sub Form_Load()
Dim ctl As Control

For Each ctl In Me.Controls
    Select Case Left(ctl.Name, 3)
        Case "chk"
            ctl.Value = -1
        
        Case "lbl"
            ctl.Caption = "- * - * -"

        Case "txt"
            ctl.Visible = False
            ctl.Value = ""
        
        Case "cmb"
            ctl.Visible = False
            
    End Select
Next ctl

Me.lstResults.RowSource = "SELECT CodMedia, Type, Titre, Boite, Pos1, Pos2, Auteur, Résumé FROM Medias;"
Me.lstResults.Requery

End Sub

Private Sub RefreshQuery()
Dim SQL As String
Dim SQLWhere As String

SQL = "SELECT CodMedia, Type, Titre, Boite, Pos1, Pos2, Auteur, Résumé FROM Medias Where Medias!CodMedia <> 0 "

If Not Me.chkAuteur Then
    SQL = SQL & "And Medias!Auteur like '*" & Me.txtRechAuteur & "*' "
End If
If Not Me.chkFamille Then
    SQL = SQL & "And Medias!Famille = '" & Me.cmbRechFamille & "' "
End If
If Not Me.chkResume Then
    SQL = SQL & "And Medias!Résumé like '*" & Me.txtRechResume & "*' "
End If
If Not Me.chkTitre Then
    SQL = SQL & "And Medias!Titre like '*" & Me.txtRechTitre & "*' "
End If
If Not Me.chkType Then
    SQL = SQL & "And Medias!Type = '" & Me.cmbRechType & "' "
End If

SQLWhere = Trim(Right(SQL, Len(SQL) - InStr(SQL, "Where ") - Len("Where ") + 1))

SQL = SQL & ";"

Me.lblStats.Caption = DCount("*", "Medias", SQLWhere) & " / " & DCount("*", "Medias")
Me.lstResults.RowSource = SQL
Me.lstResults.Requery

End Sub


Private Sub lstResults_DblClick(Cancel As Integer)

DoCmd.OpenForm "frmAutoMedias", acNormal, , "[CodMedia] = " & Me.lstResults

End Sub

Private Sub Sequence_BeforeUpdate(Cancel As Integer)
RefreshQuery
End Sub



Private Sub txtRechAuteur_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub txtRechResume_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub txtRechTitre_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub
