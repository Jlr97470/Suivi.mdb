Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
'Lors de la mise en forme des enregistrements, affiche les intitulés des champs de
'résolution uniquement pour les problèmes résolus.

    Dim blnResolved As Boolean

    'Détermine si le champ DateRésolution contient une valeur.
    blnResolved = Not IsNull(DateRésolution)

    'Positionne la propriété Visible pour les 3 intitulés.
    lbDateRésolution.Visible = blnResolved
    lbRésolution.Visible = blnResolved
    lbRésoluPar.Visible = blnResolved


End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
'Alterne les arrière-plans des sections avec blanc et gris clair.

    Const conLtGray = 13816530

    If Detail.BackColor = conLtGray Then
        Detail.BackColor = vbWhite
    Else
        Detail.BackColor = conLtGray
    End If

    'Incrémente les champs de totaux si le problème est en cours ou de priorité haute.
    If PrintCount = 1 Then
        If Etat = "En cours" Then
            ProblèmesEnCours = ProblèmesEnCours + 1
        End If
        If Priorité = "Haute" Then
            PrioritésHautes = PrioritésHautes + 1
        End If
    End If

End Sub

Private Sub EntêteGroupe0_Print(Cancel As Integer, PrintCount As Integer)
'Réinitialise les totaux dans le pied de groupe de chaque employé.

    ProblèmesEnCours = 0
    PrioritésHautes = 0

End Sub
