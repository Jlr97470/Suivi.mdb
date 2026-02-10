Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub AssignéA_AfterUpdate()
'Interroge la zone de liste pour n'afficher que les problèmes du l'employé sélectionné.

    ListeProblème.Requery

End Sub

Private Sub Form_Open(Cancel As Integer)
'Donne la valeur de l'employé en cours au champ AssignéA.

    AssignéA = lngRéfEmployéEnCours
    ListeProblème.Requery

End Sub

Private Sub ToutRéassigner_Click()
'Construit une requête d'action pour réassigner les problèmes.

    Dim strMessage As String
    Dim cnn As Connection

    'Vérifie que des valeurs sont sélectionnées dans les deux listes.
    If IsNull(AssignéA) Or IsNull(NouvelleAssignation) Then
        DisplayMessage "Vous devez sélectionner des employés " & _
            "à qui et de qui assigner des problèmes."
        Exit Sub
    End If

    'Confirme la réassignation.
    strMessage = "Réassigner tous les problèmes de la liste à " & _
        NouvelleAssignation.Column(1) & "?"
    If Confirm(strMessage) Then

        'Affiche l'icône du sablier pendant la réassignation des problèmes.
        DoCmd.Hourglass True

        'Exécute une requête UPDATE pour réassigner les problèmes.
        Set cnn = CurrentProject.Connection
        cnn.Execute "UPDATE Suivi SET AssignéA = " & _
            NouvelleAssignation & " WHERE AssignéA = " & AssignéA

        'Actualise le formulaire pour afficher les problèmes réassignés.
        AssignéA = NouvelleAssignation
        NouvelleAssignation = Null
        ListeProblème.Requery
        DoCmd.Hourglass False

    End If

End Sub

Sub FermerFormulaire_Click()
On Error GoTo Err_FermerFormulaire_Click


    DoCmd.Close acForm, "Réassigner"

Exit_FermerFormulaire_Click:
    Exit Sub

Err_FermerFormulaire_Click:
    MsgBox Err.Description
    Resume Exit_FermerFormulaire_Click

End Sub

Private Sub NouvelleAssignationSélectionnée_Click()
'Ouvre un jeu d'enregistrements à réassigne les problèmes sélectionnés.

    Dim strMessage As String
    Dim cnn As Connection
    Dim rstSuivi As New ADODB.Recordset
    Dim varPosition As Variant

    'Vérifie qu'une valeur est sélectionnée dans le champ NouvelleAssignation.
    If IsNull(NouvelleAssignation) Then
        DisplayMessage "Vous devez sélectionner un employé pour lui assigner des problèmes."
        Exit Sub
    End If

    'Vérifie qu'au moins un problème est sélectionné.
    If ListeProblème.ItemsSelected.Count = 0 Then
        DisplayMessage "Vous devez sélectionner un problème à réassigner."
        Exit Sub
    End If

    'Confirme la nouvelle assignation.
    strMessage = "Réassigner le(s) problème(s) sélectionné(s) à " & NouvelleAssignation.Column(1) & "?"
    If Confirm(strMessage) Then

        'Affiche l'icône du sablier pour la réassignation des problèmes.
        DoCmd.Hourglass True

        'Ouvre un objet recordset dans la base de données en cours.
        Set cnn = CurrentProject.Connection
        rstSuivi.Open "Suivi", cnn, adOpenKeyset, adLockOptimistic, _
            adCmdTableDirect

        'Démarre une transaction avant de modifier les données.
        cnn.BeginTrans

        'Définit la propriété Index pour rechercher la clé principale.
        rstSuivi.Index = "RéfProblème"

        'Boucle sur chaque problème sélectionné dans la liste.
        For Each varPosition In ListeProblème.ItemsSelected

            'Chercher l'enregistrement dans la table Suivi.
            rstSuivi.Seek ListeProblème.ItemData(varPosition)

            'Remplace la valeur AssignéA dans la table.
            rstSuivi!AssignéA = NouvelleAssignation
            rstSuivi.Update

        Next varPosition

        'Enregistre toutes les modifications de cette transaction.
        cnn.CommitTrans

        'Actualise le formulaire pour afficher les problèmes réassignés.
        AssignéA = NouvelleAssignation
        NouvelleAssignation = Null
        ListeProblème.Requery
        DoCmd.Hourglass False

    End If

End Sub
