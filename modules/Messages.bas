Attribute VB_Name = "Messages"
Option Compare Database
Option Explicit

Public Const conNomApp = "Suivi"

Public Sub DisplayMessage(strMessage As String)
'Affiche un message important à l'intention de l'utilisateur.

    MsgBox strMessage, vbExclamation, conNomApp

End Sub

Public Function Confirm(strMessage As String) As Boolean
'Demande à l'utilisateur de confirmer une action, returne True ou False.

    Dim bytChoice As Byte

    bytChoice = MsgBox(strMessage, vbQuestion + vbOKCancel, conNomApp)

    If bytChoice = vbOK Then
        Confirm = True
    Else
        Confirm = False
    End If

End Function
