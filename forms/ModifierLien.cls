Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Annuler_Click()
On Error GoTo Err_Annuler_Click


    DoCmd.Close

Exit_Annuler_Click:
    Exit Sub

Err_Annuler_Click:
    MsgBox Err.Description
    Resume Exit_Annuler_Click

End Sub
