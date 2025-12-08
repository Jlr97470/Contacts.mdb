Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Report_NoData(Cancel As Integer)
    MsgBox "Il n'existe aucune donnée pour cet état. Il va être annulé..."
    Cancel = -1
End Sub
Private Sub Report_Close()
    DoCmd.Close acForm, "Période de l'état"
End Sub
Private Sub Report_Open(Cancel As Integer)
    DoCmd.OpenForm "Période de l'état", , , , , acDialog, "Statistiques appels hebdo"
    If Not IsLoaded("Période de l'état") Then
        Cancel = True
    End If
End Sub
