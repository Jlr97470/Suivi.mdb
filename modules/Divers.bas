Attribute VB_Name = "Divers"
Option Compare Database
Option Explicit

Public lngRéfEmployéEnCours As Long

Public Function CréerMémoWord()
'Ouvre un mémo dans Word et insère le texte - utilisé par la commande de menu.

    Dim rstEmployés As New ADODB.Recordset
    Dim appWord As New Word.Application
    Dim intPages As Integer, strMessage As String

    'Ouvre un jeu d'enregistrements basé sur la requête EmployésAvecProblèmesNonRésolus.
    rstEmployés.Open "EmployésAvecProblèmesNonRésolus", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    'Si personne n'a de problème en cours, affiche un message et quitte.
    If rstEmployés.RecordCount = 0 Then
        DisplayMessage "Aucun problème en cours."
        Exit Function
    End If

    'Ouvre un document basé sur l'exemple Mémo, désactive la vérification orthographique,
    'se place sur le signet LigneAMémo puis affiche Word.
    With appWord
        .Documents.Add "C:\Fichiers exemples Access VB\Ch14 Mémo.dot"
        .ActiveDocument.ShowSpellingErrors = False
        .Selection.GoTo wdGoToBookmark, Name:="LigneAMémo"
        .Visible = True
    End With

    'Boucle dans le jeu d'enregistrements retourné par la requête, en insérant
    'le nom de chaque employé dans le document.
    Do Until rstEmployés.EOF
        appWord.Selection.TypeText rstEmployés!NomEmployé & "  "
        rstEmployés.MoveNext
    Loop

    'Retourne dans Access.
    AppActivate "Suivi"

    'Trouve le nombre de pages du mémo, demande à l'utilisateur s'il faut l'imprimer,
    'puis indique à Word de l'imprimer.
    intPages = appWord.Selection.Information(wdNumberOfPagesInDocument)
    strMessage = "Le mémo est terminé et contient " & CStr(intPages) & _
        " page(s). Voulez-vous l'envoyer à l'imprimante ?"
    If Confirm(strMessage) Then
        appWord.ActiveDocument.PrintOut
    End If

End Function

Public Sub OuvrirContactOutlook(strFullName As String)
'Ouvre un contact Microsoft Outlook pour le nom spécifié.

    Dim appOutlook As New Outlook.Application
    Dim nsOutlook As NameSpace
    Dim mfContacts As MAPIFolder
    Dim ciContact As ContactItem

    'Ouvre le dossier Contacts par défaut d'Outlook.
    Set nsOutlook = appOutlook.GetNamespace("MAPI")
    Set mfContacts = nsOutlook.GetDefaultFolder(olFolderContacts)

    'Cherche le contact appelé dans la colonne AssignéA.
    Set ciContact = mfContacts.Items.Find("[FullName] = " & strFullName)

    'Si le contact n'existe pas, crée un nouveau contact.
    If ciContact Is Nothing Then
        Set ciContact = mfContacts.Items.Add
        ciContact.FullName = strFullName
    End If

    'Ouvre la fenêtre Contact d'Outlook.
    ciContact.Display

End Sub


Public Function ApercuEtatSuivi()
'Ouvre l'état problèmes - utilisé par une commande de menu.

    DoCmd.OpenReport "Suivi", acViewPreview

End Function

Sub ModifierCouleurEtiquettes(strFormName As String, lngColorCode As Long)
'Modifie la couleur des étiquettes du formulaire spécifié

    Dim ctl As Control

    For Each ctl In Forms(strFormName).Controls
        If TypeOf ctl Is Label Then
            ctl.ForeColor = lngColorCode
        End If
    Next ctl
End Sub
