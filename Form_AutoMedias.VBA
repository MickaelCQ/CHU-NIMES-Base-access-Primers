Option Compare Database

' Banque de Primers
Private Sub Ajouter_Exit(Cancel As Integer)
Dim lgDernBoite As Long
Dim lgDernPos As Long
Dim strSQL As String
Dim strMsg As String
' Recherche de la dernière boite
lgDernBoite = Nz(DMax("Boite", "Medias"), 0)
 
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
 If Me.Type.Value = "KCNQ2" Then
    Me.Pathologie.Value = "SLA"
    Me.Nb_exons.Value = "17"
    Me.Num_chromosome.Value = "20"
    Me.Nom_prot.Value = "Kvoltage-gated channel subfamily Q member 2"
    Me.Editeur.Value = "NM_172107.4"
    
    ElseIf Me.Type.Value = "LDLR" Then
    Me.Pathologie.Value = "HF"
    Me.Nb_exons.Value = "18"
    Me.Num_chromosome.Value = "19"
    Me.Nom_prot.Value = "Low Density Lipoprotein Receptor"
    Me.Editeur.Value = "NM_000527.4"

    ElseIf Me.Type.Value = "APOB" Then
    Me.Pathologie.Value = "HF"
    Me.Nb_exons.Value = "29"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Apoliprotein B"
    Me.Editeur.Value = "NM_000384.3"
    
    ElseIf Me.Type.Value = "PCSK9" Then
    Me.Pathologie.Value = "HF"
    Me.Nb_exons.Value = "12"
    Me.Num_chromosome.Value = "1"
    Me.Nom_prot.Value = "Proprotein convertase subtilisine/Kexine T9"
    
    ElseIf Me.Type.Value = "LDLRAP1" Then
    Me.Pathologie.Value = "HF"
    Me.Nb_exons.Value = "9"
    Me.Num_chromosome.Value = "1"
    Me.Nom_prot.Value = "Low Density Lipoprotein Receptor Adaptator P1"

    ElseIf Me.Type.Value = "DYNC1H1" Then
    Me.Pathologie.Value = "TSA"
    Me.Nb_exons.Value = "78"
    Me.Num_chromosome.Value = "14"
    Me.Nom_prot.Value = "Dynein, cytoplasmic 1, heavy chain 1"
    Me.Editeur.Value = "NM_001376.5"
    
    ElseIf Me.Type.Value = "APOE" Then
    Me.Pathologie.Value = "HF"
    Me.Nb_exons.Value = "4"
    Me.Num_chromosome.Value = "19"
    Me.Nom_prot.Value = "Apolipoprotein E"""
    
    ElseIf Me.Type.Value = "ALS2" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "34"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Amyotrophic lateral sclerosis 2"
    
    ElseIf Me.Type.Value = "POGZ" Then
    Me.Pathologie.Value = "TSA"
    Me.Nb_exons.Value = "19"
    Me.Num_chromosome.Value = "1"
    Me.Nom_prot.Value = "Pogo transposable element with ZNF domain"
    Me.Editeur.Value = "NM_015100.4"
    
    ElseIf Me.Type.Value = "ARID1B" Then
    Me.Pathologie.Value = "TSA"
    Me.Nb_exons.Value = "20"
    Me.Num_chromosome.Value = "6"
    Me.Nom_prot.Value = "AT-rich interaction domain 1B"
    Me.Editeur.Value = "NM_001374820.1"
    
    ElseIf Me.Type.Value = "ANG" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "2"
    Me.Num_chromosome.Value = "14"
    Me.Nom_prot.Value = "Angiogenin"
    
    ElseIf Me.Type.Value = "APP" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "18"
    Me.Num_chromosome.Value = "21"
    Me.Nom_prot.Value = "Amyloid precursor protein"
    Me.Editeur.Value = "NM_000484.4"
    
    ElseIf Me.Type.Value = "C9ORF72" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "11"
    Me.Num_chromosome.Value = "9"
    Me.Nom_prot.Value = "Chromosome 9 open reading frame 72"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "C21ORF2" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "0"
    Me.Num_chromosome.Value = "0"
    Me.Nom_prot.Value = "non indiqué"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CAPRIN" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "19"
    Me.Num_chromosome.Value = "11"
    Me.Nom_prot.Value = "Cell cycle associated protein 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CCNF" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "17"
    Me.Num_chromosome.Value = "16"
    Me.Nom_prot.Value = "Cyclin F"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CHCHD10" Then
    Me.Pathologie.Value = "SLA"
    Me.Nb_exons.Value = "4"
    Me.Num_chromosome.Value = "22"
    Me.Nom_prot.Value = "coiled coil helix-[...] domain containing 10"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CHMP2B" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "6"
    Me.Num_chromosome.Value = "3"
    Me.Nom_prot.Value = "Charged multivesicular body protein 2B"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CSF1R" Then
    Me.Pathologie.Value = "Leucodystrophie"
    Me.Nb_exons.Value = "22"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "Colony stimulating factor 1 receptor"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DAO" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "11"
    Me.Num_chromosome.Value = "12"
    Me.Nom_prot.Value = "D-amino-acid oxidase"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DCTN1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "32"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Dynactin 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DPYSL3" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "14"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "dihydropyrimidinase like 3"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "ELP3" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "15"
    Me.Num_chromosome.Value = "8"
    Me.Nom_prot.Value = "Elongator acetyltransferase complex subunit 3"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "EPHA4" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "19"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "EPH receptor A4"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "TRIP12" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "42"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Thyroid Hormone receptor interactor 12 "
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CCNF" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "17"
    Me.Num_chromosome.Value = "16"
    Me.Nom_prot.Value = "Cyclin F "
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CHMP2B" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "6"
    Me.Num_chromosome.Value = "3"
    Me.Nom_prot.Value = "Charged multivesicular body protein 2B"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "CSF1R" Then
    Me.Pathologie.Value = "Leucodystrophie"
    Me.Nb_exons.Value = "6"
    Me.Num_chromosome.Value = "3"
    Me.Nom_prot.Value = "Charged multivesicular body protein 2B"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DAO" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "11"
    Me.Num_chromosome.Value = "12"
    Me.Nom_prot.Value = "D-amino-acid oxidase"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DCTN1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "32"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Dynactin 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "DPYSL3" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "14"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "dihydropyrimidinase like 3"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "ELP3" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "15"
    Me.Num_chromosome.Value = "8"
    Me.Nom_prot.Value = "Elongator acetyltransferase complex subunit 3"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "EPHA4" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "19"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "EPH receptor A4"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "ERBB4" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "28"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Erb-b2 receptor tyrosine kinase 4"
    Me.Editeur.Value = "NM_005235.3"
    
    ElseIf Me.Type.Value = "EWSR1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "18"
    Me.Num_chromosome.Value = "22"
    Me.Nom_prot.Value = "EWS RNA-binding protein 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = ""
    Me.Num_chromosome.Value = ""
    Me.Nom_prot.Value = ""
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "FIG4" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "23"
    Me.Num_chromosome.Value = "6"
    Me.Nom_prot.Value = "FIG4 Phosphoinositide 5-phosphatase"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "FUS" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "15"
    Me.Num_chromosome.Value = "16"
    Me.Nom_prot.Value = "FUS RNA binding protein"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = ""
    Me.Num_chromosome.Value = ""
    Me.Nom_prot.Value = ""
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "GRN" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "13"
    Me.Num_chromosome.Value = "17"
    Me.Nom_prot.Value = "Granulin"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "KIF5A" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "29"
    Me.Num_chromosome.Value = "12"
    Me.Nom_prot.Value = "Kinesin family member 5A"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "HNRNPA1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "11"
    Me.Num_chromosome.Value = "12"
    Me.Nom_prot.Value = "Heterogeneous nuclear ribonucleoprotein A1"
    Me.Editeur.Value = "?"
    
     ElseIf Me.Type.Value = "HNRNPA2B1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "12"
    Me.Num_chromosome.Value = "7"
    Me.Nom_prot.Value = "Heterogen nucleat ribonucleoprotein A2/B1"
    Me.Editeur.Value = "?"
    
     ElseIf Me.Type.Value = "LMNB1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "11"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "Lamin B1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "MAPT" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "15"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "Microtubule associated protein tau"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "MATR3" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "18"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "Matrin 3"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "NEFH" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "4"
    Me.Num_chromosome.Value = "22"
    Me.Nom_prot.Value = "Neurofilament, heavy polypeptide"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "NEK1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "36"
    Me.Num_chromosome.Value = "4"
    Me.Nom_prot.Value = "NIMA-related kinase 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "OPTN" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "16"
    Me.Num_chromosome.Value = "10"
    Me.Nom_prot.Value = "Optineurin"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "PFN1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "3"
    Me.Num_chromosome.Value = "17"
    Me.Nom_prot.Value = "Profilin 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "PRPH" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "12"
    Me.Num_chromosome.Value = "9"
    Me.Nom_prot.Value = "Peripherin"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "PSEN1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "12"
    Me.Num_chromosome.Value = "14"
    Me.Nom_prot.Value = "Prenisilin 1"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "PSEN2" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "13"
    Me.Num_chromosome.Value = "1"
    Me.Nom_prot.Value = "Prenisilin 2"
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "SETX" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "27"
    Me.Num_chromosome.Value = "9"
    Me.Nom_prot.Value = "Senataxin"
    Me.Editeur.Value = "NM_015046"
    
    ElseIf Me.Type.Value = "SIGMAR1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "4"
    Me.Num_chromosome.Value = "9"
    Me.Nom_prot.Value = "Sigma non-opioid intracellular receptor 1"
    Me.Editeur.Value = "A renseigner"
    
    ElseIf Me.Type.Value = "SOD1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "5"
    Me.Num_chromosome.Value = "21"
    Me.Nom_prot.Value = "Superoxide dismutase 1, soluble"
    Me.Editeur.Value = "ENST00000270142?"
    
    ElseIf Me.Type.Value = "SPAST" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "17"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Spastin"
    Me.Editeur.Value = "NM_014946"
    
    ElseIf Me.Type.Value = "SPG11" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "40"
    Me.Num_chromosome.Value = "15"
    Me.Nom_prot.Value = "Spastic paraplegia 11"
    Me.Editeur.Value = "NM_025137"
    
    ElseIf Me.Type.Value = "SPTLC1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "15"
    Me.Num_chromosome.Value = "9"
    Me.Nom_prot.Value = "Serie palmitoyltransferase, long chain base subunit  1"
    Me.Editeur.Value = "A renseigner"
    
    ElseIf Me.Type.Value = "SQSTM1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "9"
    Me.Num_chromosome.Value = "5"
    Me.Nom_prot.Value = "Sequestosome 1"
    Me.Editeur.Value = "A renseigner"
    
    ElseIf Me.Type.Value = "TAF15" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "16"
    Me.Num_chromosome.Value = "17"
    Me.Nom_prot.Value = "TAF15 RNA Pol II,(TBP)-associated factor"
    Me.Editeur.Value = "NM_139215"
    
    ElseIf Me.Type.Value = "TARDBP" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "6"
    Me.Num_chromosome.Value = "1"
    Me.Nom_prot.Value = "TAR DNA binding protein"
    Me.Editeur.Value = "A renseigner"
    
    ElseIf Me.Type.Value = "TBK1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "21"
    Me.Num_chromosome.Value = "12"
    Me.Nom_prot.Value = "Tank binding kinase 1"
    Me.Editeur.Value = "NM_013254"
    
    ElseIf Me.Type.Value = "TIA1" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "13"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "cytotoxic granule associated RNA binding protein"
    Me.Editeur.Value = "NM_022173.4"
    
    ElseIf Me.Type.Value = "TREM2" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "5"
    Me.Num_chromosome.Value = "6"
    Me.Nom_prot.Value = "Triggering receptor expressed on myeloid cells 2"
    Me.Editeur.Value = "A renseigner"
    
    ElseIf Me.Type.Value = "TUBA4A" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = "4"
    Me.Num_chromosome.Value = "2"
    Me.Nom_prot.Value = "Tubulin, alpha 4a"
    Me.Editeur.Value = "NM_006000.3"
    
    ElseIf Me.Type.Value = "" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = ""
    Me.Num_chromosome.Value = ""
    Me.Nom_prot.Value = ""
    Me.Editeur.Value = "?"
    
    ElseIf Me.Type.Value = "" Then
    Me.Pathologie.Value = "SLA/DFT"
    Me.Nb_exons.Value = ""
    Me.Num_chromosome.Value = ""
    Me.Nom_prot.Value = ""
    Me.Editeur.Value = "?"
    
  
End If
End Sub
