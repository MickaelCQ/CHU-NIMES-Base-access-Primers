Option Compare Database

Private Sub Btn1_Ouvrir_Gesqual_Click()
Dim strString As String
strString = "http://192.168.21.245/"
Application.FollowHyperlink strString
End Sub

Private Sub chkNgmao_Click()

If Me.chkNgmao Then
    Me.txtRechNgmao.Visible = False
Else
    Me.txtRechNgmao.Visible = True
End If

RefreshQuery

End Sub


Private Sub chkDiscipline_Click()

If Me.chkDiscipline Then
    Me.cmbRechDiscipline.Visible = False
Else
    Me.cmbRechDiscipline.Visible = True
End If

RefreshQuery

End Sub

Private Sub chkNserie_Click()

If Me.chkNserie Then
    Me.txtRechNserie.Visible = False
Else
    Me.txtRechNserie.Visible = True
End If

RefreshQuery

End Sub

Private Sub chkNom_Click()

If Me.chkNom Then
    Me.txtRechNom.Visible = False
Else
    Me.txtRechNom.Visible = True
End If

RefreshQuery

End Sub


Private Sub chkCategorie_Click()

If Me.chkCategorie Then
    Me.cmbRechCategorie.Visible = False
Else
    Me.cmbRechCategorie.Visible = True
End If

RefreshQuery

End Sub

Private Sub cmbRechDiscipline_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub cmbRechCategorie_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub


Private Sub cmbRechCategorie_Click()

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

Me.lstResults2.RowSource = "SELECT CodMateriel, Categorie, Nom, piece, critique, statut, Ngmao, Nserie FROM Materiels;"
Me.lstResults2.Requery

End Sub

Private Sub RefreshQuery()
Dim SQL As String
Dim SQLWhere As String

SQL = "SELECT CodMateriel, Categorie, Nom, Ngmao, piece, critique, statut, Nserie FROM Materiels Where Materiels!CodMateriel <> 0 "

If Not Me.chkNgmao Then
    SQL = SQL & "And Materiels!Ngmao like '*" & Me.txtRechNgmao & "*' "
End If
If Not Me.chkDiscipline Then
    SQL = SQL & "And Materiels!Discipline = '" & Me.cmbRechDiscipline & "' "
End If
If Not Me.chkNserie Then
    SQL = SQL & "And Materiels!Nserie like '*" & Me.txtRechNserie & "*' "
End If
If Not Me.chkNom Then
    SQL = SQL & "And Materiels!Nom like '*" & Me.txtRechNom & "*' "
End If
If Not Me.chkCategorie Then
    SQL = SQL & "And Materiels!Categorie = '" & Me.cmbRechCategorie & "' "
End If

SQLWhere = Trim(Right(SQL, Len(SQL) - InStr(SQL, "Where ") - Len("Where ") + 1))

SQL = SQL & ";"

Me.lblstats2.Caption = DCount("*", "Materiels", SQLWhere) & " / " & DCount("*", "Materiels")
Me.lstResults2.RowSource = SQL
Me.lstResults2.Requery

End Sub


Private Sub Genes_AfterUpdate()
Sequences_Word = Genes.Column(2)
Pathologie_associee = Genes.Column(3)
Pathologie_Nom = Genes.Column(4)

End Sub

Private Sub lstResults2_DblClick(Cancel As Integer)

DoCmd.OpenForm "frmAutoMateriels", acNormal, , "[CodMateriel] = " & Me.lstResults2

End Sub

Private Sub Sequence_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub txtRechNgmao_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub txtRechNserie_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub

Private Sub txtRechNom_BeforeUpdate(Cancel As Integer)

RefreshQuery

End Sub


