Option Compare Database

' Banque de Primers
Private Sub Ajouter_Exit(Cancel As Integer)
Dim lgDernBoite As Long
Dim lgDernPos As Long
Dim strSQL As String
Dim strMsg As String
' Recherche de la dernière boite
lgDernBoite = Nz(DMax("Boite", "Medias"), 0)


' A l'intention des tech : Ne jamais utiliser la position 80 , sinon décalage dans les attributions et BUG du code pour le saut de boite. Mickael 26/05/2024.

' Recherche dernière position de la dernière boite

If lgDernBoite <> 0 Then
   lgDernPos = Nz(DMax("pos2", "Medias", "Boite=" & lgDernBoite), 0)
' chargement d'une nouvelle boite si on a atteint les 80 positions
    If lgDernPos = 80 Then
        lgDernBoite = lgDernBoite + 1
        lgDernPos = 1
        strMsg = "la dernière boite est pleine, vous êtes sur le point de créer une nouvelle  ?"
    Else
' sinon on ajoute 2 positions à la boite existante
        strMsg = "Voulez-vous créer deux nouvelles positions pour la boîte actuelle ?"
        lgDernPos = lgDernPos + 1
    End If
Else
    lgDernBoite = 1
    lgDernPos = 1
    strMsg = "Voulez-vous créer la première boîte ?"
End If
 
If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
    strSQL = "Insert into Medias(Boite,Pos1,Pos2) values (" & lgDernBoite & "," & lgDernPos & "," & lgDernPos + 1 & ");"
' Exécution de la requète d'ajout
    CurrentDb.Execute strSQL, dbFailOnError
' mise à jour du formulaire et affichage du dernier enregistrement
    Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End If

End Sub

Private Sub Aide_intitule_Primers_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Il faut nommer les Primers exactement comme intitulé sur les flacons de solution mère "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Aide dénomination Primers "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub


Private Sub Aide_sequence_Primers_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "M13 Fw: TGTAAAACGACGGCCAGT Rv:GGATAACAATTTCACACAGG"
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Aide séquences M13 "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub


Private Sub Aide_Sequences_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Les positions dans la banque s'incrémentent automatiquement. Néanmoins, à chaque sur : 'Saisir un nouveau couple de Primers', une nouvelle ligne d'écriture s'inscrit dans la base de données. si vous n'allez pas au bout de la saisie ou dans le cas d'une fausse manipulation prévenir Mickael pour purger la base et éviter que des positions dans les boites restent vides."
Style = vbYes + vbCritical + vbDefaultButton2
Title = "Aide séquences Primers "
Ctxt = 1000
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then
    MyString = "Fermer"
End If
End Sub

Private Sub Ajout_reception_1_Click()
If Me.Date_BL_1.Value = True Then
Date_BL_1.Visible = True
BL_1.Visible = True
Supr_reception_1.Visible = True
ElseIf Me.Date_BL_1.Value = False Then
Date_BL_1.Visible = False
BL_1.Visible = False
Supr_reception_1.Visible = False
End If
End Sub
Private Sub Ajout_reception_2_Click()
Date_BL_2.Visible = True
BL_2.Visible = True
Supr_reception_2.Visible = True
End Sub
Private Sub Ajout_reception_3_Click()
Date_BL_3.Visible = True
BL_3.Visible = True
Supr_reception_3.Visible = True
End Sub
Private Sub Ajout_reception_4_Click()
Date_BL_4.Visible = True
BL_4.Visible = True
Supr_reception_4.Visible = True
End Sub

Private Sub Bouton_Design_primer_ok_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Phase 1 : Couple de Primers dessiné et intégré dans la fiche d'informations "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information sur le Design des primers"    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Bouton_Commande_primer_ok_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Phase 2 : Commande du couple de Primers effectuée, en attente de réception. "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information sur le dessin des primers "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Bouton_Reception_primer_ok_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Phase 3 : Primers récéptionnés, la commande a été vérifiée et stockée "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information sur la réception des primers "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Bouton_test_Primers_ok_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Phase 4 : La présence de cette cartouche signale que les Primers sont en routine depuis plus de 6 mois "    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information sur l'utilisation "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Bouton_test_Primers_pas_OK2_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Phase 4 : Cette cartouche signale que le couple d'amorces a été testé : avec les Taq Promega et Robust, En cochant la case Echec PCR pensez à : redessiner et recreer un nouvel enregistrement"    ' Define message.
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Information sur l'échec de PCR"    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If
End Sub

Private Sub Date_BL_1_AfterUpdate()
Me!BL_1 = " 1er reception cliquez-ici " & "#B:\ROSALIND\2-Primers\RECEPTIONS_PRIMERS\" & Date_BL_1 & ".pdf" & "#"
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Date_BL_2_AfterUpdate()
Me!BL_2 = "2e reception cliquez-ici " & "#B:\ROSALIND\2-Primers\RECEPTIONS_PRIMERS\" & Date_BL_2 & ".pdf" & "#"
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Date_BL_3_AfterUpdate()
Me!BL_3 = "3e reception cliquez-ici " & "#B:\ROSALIND\2-Primers\RECEPTIONS_PRIMERS\" & Date_BL_3 & ".pdf" & "#"
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Date_BL_4_AfterUpdate()
Me!BL_4 = "4e reception cliquez-ici " & "#B:\ROSALIND\2-Primers\RECEPTIONS_PRIMERS\" & Date_BL_4 & ".pdf" & "#"
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Docs_annexes_Click()
Dim Msg, Style, Title, Ctxt, Response, MyString
Msg = "Pour consulter un bon de livraison, cliquez sur (+) et renseignez  une date de reception au format JJ.MM.AAAA , l'informatique s'occupe du reste !"
Style = vbYes + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "Comment ajouter un bon de livraison ? "    ' Define title.
Ctxt = 1000    ' Define topic context.
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Fermer"    ' Perform some action.
End If

End Sub

' affichage du contrôl PopUpRenouvellement
Private Sub Form_Current()
If [Compteur_jour_renouvellement] <= 0 Then
[PopUpRenouvellement].Visible = True
ElseIf [Compteur_jour_renouvellement] > 0 Then
[PopUpRenouvellement].Visible = False
End If

'Conditions affichage cartouche " Echec PCR"
  If Me.Chck_echec_PCR.Value = True Then
    Me.Bouton_test_Primers_pas_OK.Visible = True
    ElseIf Me.Chck_echec_PCR.Value = False Then
    Me.Bouton_test_Primers_pas_OK.Visible = False
    End If
    
    If Me.Chck_echec_PCR.Value = True Then
    Me.Bouton_chevron_4_Echec.Visible = True
    ElseIf Me.Chck_echec_PCR.Value = False Then
    Me.Bouton_chevron_4_Echec.Visible = False
    End If
    
     If Me.Chck_echec_PCR.Value = True Then
    Me.Bouton_test_Primers_pas_OK2.Visible = True
    ElseIf Me.Chck_echec_PCR.Value = False Then
    Me.Bouton_test_Primers_pas_OK2.Visible = False
    End If
    
    If Me.Chck_echec_PCR.Value = True Then
    Me.Bouton_chevron_5_Echec.Visible = True
    ElseIf Me.Chck_echec_PCR.Value = False Then
    Me.Bouton_chevron_5_Echec.Visible = False
    End If

'Conditions affichage cartouche "Primer designé" et Chevron1 associé

If Not IsNull(Me.Sequence_Primer_F) And Not IsNull(Me.Sequence_Primer_R) Then
    Bouton_Design_primer_ok.Visible = True
    Else
    Bouton_Design_primer_ok.Visible = False
End If

If Bouton_Design_primer_ok.Visible = True Then
    Bouton_chevron_1.Visible = True
    Else
    Bouton_chevron_1.Visible = False
End If

'Conditions affichage cartouche "Primer commandé" Chevron2 associé
If Me.Chck_Commande_primers.Value = False Then
    Me.Bouton_Commande_primer_ok.Visible = False
    
    ElseIf Me.Chck_Commande_primers.Value = True Then
    Me.Bouton_Commande_primer_ok.Visible = True
    End If
    
 If Me.Chck_Commande_primers.Value = True Then
    Bouton_chevron_2.Visible = True
    Else
    Bouton_chevron_2.Visible = False
End If

'Conditions affichage cartouche "Primer réceptionné" Chevron3 associé
If Not IsNull(Me.Date_reception) Then
    Bouton_Reception_primer_ok.Visible = True
    Else
    Bouton_Reception_primer_ok.Visible = False
    End If
    
 If Bouton_Reception_primer_ok.Visible = True Then
    Bouton_chevron_3.Visible = True
    Else
    Bouton_chevron_3.Visible = False
End If

'Conditions affichage cartouche "Primers Testés - PCR OK"
If [Compteur_jour_reception] > 183 Then
    Bouton_test_Primers_ok.Visible = True
    Else
    Bouton_test_Primers_ok.Visible = False
    End If
    End Sub
    
'Actualisation automatique des conditions pour MAJ de la barre de processus

Private Sub Chck_echec_PCR_AfterUpdate()
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub
Private Sub Sequence_Primer_F_AfterUpdate()
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Sequence_Primer_R_AfterUpdate()
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub

Private Sub Chck_Commande_primers_AfterUpdate()
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub


' Masquer le Ruban de création en affichage utilisateur
Private Sub Form_Load()
Me.AllowEdits = False
DoCmd.ShowToolbar "Ribbon", acToolbarNo
Me.ShortcutMenu = False

End Sub

' Module d'incrémentation automatique des banques de Primers
Private Sub Increment_Click()
Dim lgDernBoite As Long
Dim lgDernPos As Long
Dim strSQL As String
Dim strMsg As String
' Recherche de la dernière boite de primers ouverture
lgDernBoite = Nz(DMax("Boite", "Medias"), 0)
 
' Recherche de la dernière position de la dernière boite créer
If lgDernBoite <> 0 Then
   lgDernPos = Nz(DMax("pos2", "Medias", "Boite=" & lgDernBoite), 0)

' création d'une nouvelle boite si on a atteint les positions
    If lgDernPos = 80 Then
        lgDernBoite = lgDernBoite + 1
        lgDernPos = 1
        strMsg = "la dernière boite est pleine, vous êtes sur le point de créer une nouvelle  ?"
    Else
' sinon on ajoute 2 positions à la boite existante
        strMsg = "Voulez-vous créer deux nouvelles positions pour la boîte actuelle ?"
        lgDernPos = lgDernPos + 1
    End If
Else
    lgDernBoite = 1
    lgDernPos = 1
    strMsg = "Voulez-vous créer la première boîte ?"
End If
 
If MsgBox(strMsg, vbQuestion + vbYesNo) = vbYes Then
    strSQL = "Insert into Medias(Boite,Pos1,Pos2) values (" & lgDernBoite & "," & lgDernPos & "," & lgDernPos + 1 & ");"
' Exécution de la requète d'ajout
    CurrentDb.Execute strSQL, dbFailOnError
' mise à jour du formulaire et affichage du dernier enregistrement
Me.Filter = ""
Me.FilterOn = False
    Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End If
End Sub

' Création d'un bouton pour verouiller/deverouiller les champs du formulaire de saisie
Private Sub Libérer_Click()
If Forms![frmAutoMedias].Form.AllowEdits = False Then
Me.AllowEdits = True
Libérer.ForeColor = RGB(0, 200, 0)  'le texte du bouton est en vert, les champs sont libérés
Else
Me.AllowEdits = False
Libérer.ForeColor = RGB(255, 0, 0)  'Le texte du bouton est en rouge, les champs sont figés
End If
DoCmd.RunCommand acCmdRefresh
End Sub

Private Sub Type_AfterUpdate()
Me!Sequence_Ref = " Séquence de référence du gène " & "#B:\ROSALIND\2-Primers\SEQUENCES_DE_REFERENCE\" & [Type] & ".rtf" & "#"
Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast
End Sub
