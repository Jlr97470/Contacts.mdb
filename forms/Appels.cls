Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Open(Cancel As Integer)
    If Not IsLoaded("Contacts") Then
        MsgBox "Ouvrez le formulaire Appels à l'aide du bouton Appels du formulaire Contacts."
        Cancel = True
    End If
End Sub
Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
    Me.Requery

Exit_Form_Activate:
    Exit Sub

Err_Form_Activate:
    MsgBox Err.Description
    Resume Exit_Form_Activate
End Sub
Private Sub Sous_formulaire_Détail_des_appels_Enter()
    If IsNull(Me![RéfAppel]) Then
      Me![Sous-formulaire Liste des appels].Form![DateAppel] = Date
      DoCmd.GoToControl "Sous-formulaire Liste des appels"
      DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
      DoCmd.GoToControl "Sous-formulaire Détail des appels"
    End If
End Sub
