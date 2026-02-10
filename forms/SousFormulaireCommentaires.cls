Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub Form_BeforeInsert(Cancel As Integer)
'Définit les valeur par défaut pour les champs Employé, Date et Etat.

    RéfEmployé = lngRéfEmployéEnCours
    DateSaisie = Date

    'Utilise la propriété Parent pour faire référence au formulaire principal.
    Etat = Me.Parent!Etat

End Sub
