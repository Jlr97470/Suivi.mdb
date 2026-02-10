Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Définit les constantes pour les valeurs du champ Résolution.
Const conEnCours = 1
Const conRésolu = 2
Const conAjourné = 4

Private Sub Commentaires_Enter()
'Passe à l'enregistrement du nouveau commentaire.

    DoCmd.GoToRecord , , acNewRec

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
'Place la valeur du champ SaisiPar sur RéfEmployé dans la variable publique.

    SaisiPar = lngRéfEmployéEnCours

End Sub

Private Sub ContactOutlook_Click()
'Ouvre un Contact Outlook pour cette assignation.

    If Not IsNull(AssignéA) Then
        OuvrirContactOutlook AssignéA.Column(1)
    End If

End Sub

Private Sub Résolution_BeforeUpdate(Cancel As Integer)
'Si l'état du problème n'est pas Résolu, le confirme.

    If Etat <> conRésolu Then
        If Confirm("Résoudre ce problème ?") Then

            'Définit la date de résolution et la valeur du champ Résolu par.
            Etat = conRésolu
            DateRésolution = Date
            RésoluPar = lngRéfEmployéEnCours

            'Assigne le problème à la personne qui l'a saisi.
            AssignéA = SaisiPar
        Else
            Cancel = True
        End If
    End If

End Sub

Private Sub Etat_AfterUpdate()
'Si le problème est résolu, se place dans le champ Résolution.

    If Etat = conRésolu Then
        Résolution.SetFocus
    End If

End Sub

Private Sub Etat_BeforeUpdate(Cancel As Integer)
'Confirme le changement d'Etat, puis définit les autres champs.

    Dim strMessage As String

    'Si la valeur n'a pas changé, quitte la procédure.
    If Etat = Etat.OldValue Then Exit Sub

    Select Case Etat
    Case conEnCours

        If Confirm("Ouvrir à nouveau ce problème ?") Then

            'Vide les champs Résolution.
            RésoluPar = Null
            DateRésolution = Null
            Résolution = Null
        Else
            Cancel = True
        End If

    Case conRésolu
        If Confirm("Résoudre ce problème ?") Then

            'Définit la date de résolution et les valeurs Résolu par.
            DateRésolution = Date
            RésoluPar = lngRéfEmployéEnCours

            'Assigne le problème à la personne qui l'a saisi.
            AssignéA = SaisiPar
        Else
            Cancel = True
        End If

    Case Else

        'Définit le message approprié.
        If Etat = conAjourné Then
            strMessage = "Ajourner ce problème ?"
        ElseIf lngRéfEmployéEnCours = SaisiPar Then
            strMessage = "Fermer ce problème ?"
        Else
            strMessage = "Vous n'avez pas saisi ce problème. Voulez-vous le fermer ?"
        End If

        If Confirm(strMessage) Then
            AssignéA = Null    'Vide le champ AssignéA.
        Else
            Cancel = True
        End If

    End Select

End Sub
