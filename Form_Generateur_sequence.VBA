Option Compare Database

Private Sub Sequence_a_convertir_AfterUpdate()

If Sequence_a_convertir <> vbNullString Then Nb_G_Sequence_a_convertir = UBound(Split(Sequence_a_convertir, "G", -1, Compare))
If Sequence_a_convertir <> vbNullString Then Nb_C_Sequence_a_convertir = UBound(Split(Sequence_a_convertir, "C", -1, Compare))

Sequence_ADNc = Replace(Sequence_a_convertir, "A", "1")

If Sequence_ADNc = Sequence_a_convertir Then

Sequence_ADNc = Replace(Sequence_ADNc, "T", "A")
Sequence_ADNc = Replace(Sequence_ADNc, "1", "T")
 Else
Sequence_ADNc = Replace(Sequence_ADNc, "T", "A")
Sequence_ADNc = Replace(Sequence_ADNc, "1", "T")
 End If
Sequence_ADNc = Replace(Sequence_ADNc, "C", "1")
Sequence_ADNc = Replace(Sequence_ADNc, "G", "C")
Sequence_ADNc = Replace(Sequence_ADNc, "1", "G")

Sequence_ARN_messager = Replace(Sequence_reverse_complement, "A", "1")

If Sequence_ARN_messager = Sequence_reverse_complement Then
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "T", "A")
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "1", "U")
 Else
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "T", "A")
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "1", "U")
 End If
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "C", "1")
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "G", "C")
Sequence_ARN_messager = Replace(Sequence_ARN_messager, "1", "G")


If Not IsNull(Sequence_a_convertir) And Not IsNull(Sequence_a_convertir) Then
    Sequence_ADNc.Visible = True
    Else
    Sequence_ADNc.Visible = False
End If

If Not IsNull(Sequence_a_convertir) And Not IsNull(Sequence_a_convertir) Then
    Sequence_reverse_complement.Visible = True
    Else
    Sequence_reverse_complement.Visible = False
End If

If Not IsNull(Sequence_a_convertir) And Not IsNull(Sequence_a_convertir) Then
    Sequence_ARN_messager.Visible = True
    Else
    Sequence_ARN_messager.Visible = False
End If


Me.Requery
    DoCmd.GoToRecord acDataForm, Me.Name, acLast

End Sub


