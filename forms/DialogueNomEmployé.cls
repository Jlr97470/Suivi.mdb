Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub OK_Click()
On Error GoTo Err_OK_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Suivi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OK_Click:
    Exit Sub

Err_OK_Click:
    MsgBox Err.Description
    Resume Exit_OK_Click

End Sub

Sub AfficherTout_Click()
On Error GoTo Err_AfficherTout_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    'Vérifie si l'utilisateur a saisi un nom.
    If IsNull(cbxRéfUtilisateur) Then
        DisplayMessage "Vous devez sélectionner ou saisir votre nom."
        cbxRéfUtilisateur.SetFocus
        Exit Sub
    End If

    'Ouvre le formulaire Suivi.
    stDocName = "Suivi"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

    'Ferme la boîte de dialogue.
    DoCmd.Close acForm, "DialogueNomEmployé"

Exit_AfficherTout_Click:
    Exit Sub

Err_AfficherTout_Click:
    MsgBox Err.Description
    Resume Exit_AfficherTout_Click

End Sub

Sub FiltrerProblèmes_Click()
On Error GoTo Err_FiltrerProblèmes_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    'Vérifie si l'utilisateur a saisi un nom.
    If IsNull(cbxRéfUtilisateur) Then
        DisplayMessage "Vous devez sélectionner ou saisir votre nom."
        cbxRéfUtilisateur.SetFocus
        Exit Sub
    End If

    stDocName = "Suivi"
    stLinkCriteria = "[AssignéA]=" & cbxRéfUtilisateur

    'Vérifie si des problèmes sont assignés à l'utilisateur. Si c'est le cas, ouvre le formulaire
    'Suivi avec un filtre. Sinon, l'ouvre sans filtre.
    If DCount("RéfProblème", "Suivi", stLinkCriteria) Then
        DoCmd.OpenForm stDocName, , , stLinkCriteria
    Else
        DisplayMessage "Aucun problème ne vous est assigné -- affichage de tous les problèmes."
        DoCmd.OpenForm stDocName
    End If

    'Ferme la boîte de dialogue.
    DoCmd.Close acForm, "DialogueNomEmployé"

Exit_FiltrerProblèmes_Click:
    Exit Sub

Err_FiltrerProblèmes_Click:
    MsgBox Err.Description
    Resume Exit_FiltrerProblèmes_Click

End Sub

Private Sub cbxRéfUtilisateur_AfterUpdate()
'Copie la référence de l'employé saisie dans une variable publique. A moins que
'cette variable ne soit déclarée dans le module Divers via le mot clé Public
'elle est disponible dans l'application.

    lngRéfEmployéEnCours = cbxRéfUtilisateur

End Sub

Private Sub cbxRéfUtilisateur_NotInList(NewData As String, Response As Integer)
'Demande à l'utilisateur s'il faut ajouter la valeur à la liste.

    Dim strMessage As String
    Dim rstEmployés As New ADODB.Recordset

    strMessage = "Voulez-vous ajouter '" & NewData & _
        "' à la liste des employés ?"

    If Confirm(strMessage) Then

        'Ouvre la table Types de contact et ajoute la valeur NewData.
        rstEmployés.Open "Employés", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rstEmployés.AddNew
        rstEmployés!NomEmployé = NewData
        rstEmployés.Update
        cbxRéfUtilisateur = rstEmployés!RéfEmployé

        Response = acDataErrAdded     'Interroge la liste.
    Else
        Response = acDataErrDisplay   'Affiche l'erreur.
    End If

End Sub
