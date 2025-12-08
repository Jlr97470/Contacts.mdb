Attribute VB_Name = "Messages"
Option Compare Database
Option Explicit

Const conNomApp = "Gestion des contacts"

Public Sub AfficheMessage(strMessage As String)
'Affiche un message important à l'intention de l'utilisateur.
    MsgBox strMessage, vbExclamation, conNomApp
End Sub

Public Function Confirmer(strMessage As String) As Boolean
    'Demande à l'utilisateur de confirmer une action en retournant True ou False.
    Dim bytChoice As Byte
    bytChoice = MsgBox(strMessage, vbQuestion + vbOKCancel, conNomApp)
    If bytChoice = vbOK Then
        Confirmer = True
    Else
        Confirmer = False
    End If

End Function
