Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CodePostal_AfterUpdate()
'Si code postal de Paris ou Lyon, complète les champs Ville et DepartementOuRegion.

    Select Case CodePostal
                Case "75000" To "75999"
            Ville = "Paris"
            DepartementOuRegion = "Paris"
        Case "69000" To "69599"
            Ville = "Lyon"
            DepartementOuRegion = "Rhône"
    End Select
    If CodePostal > "00000" And CodePostal < "99999" Then
            Pays = "France"
            Titre.SetFocus
        End If
End Sub
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    'Positionne la valeur du groupe d'options pour qu'elle corresponde à l'action de filtrage choisie par l'utilisateur.

        If ApplyType = acShowAllRecords Then
            OptionsFiltre = 1       'Active l'option Tous les contacts.
        ElseIf Filter <> "Ville = 'Paris'" Then
            OptionsFiltre = Null        'désélectionne les deux options.
        End If
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Si l'utilisateur a saisi une adresse, vérifie le code postal.
    Dim blnOK As Boolean


    If Not IsNull(Adresse) And IsNull(CodePostal) Then
          blnOK = Confirmer("Vous n'avez pas saisi de code postal. Voulez-vous valider l'enregistrement ?")

        If Not blnOK Then   'L'utilisateur a cliqué sur Annuler.
            CodePostal.SetFocus 'Retourne dans le champ CodePostal.
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Current()
    Modifiable50 = RéfContact       'Met à jour la zone de liste Trouver enregistrement.


End Sub

Private Sub Form_Filter(Cancel As Integer, FilterType As Integer)
    'Affiche instructions pour Filtrer par formulaire.
    If FilterType = acFilterByForm Then
        MsgBox "Sélectionnez des valeurs de champ et cliquez sur le bouton Appliquer le filtre."
    End If


End Sub

Private Sub Modifiable50_Enter()
    Modifiable50.BackColor = 16777215   'place l'arrière-plan en blanc

End Sub


Private Sub Modifiable50_Exit(Cancel As Integer)
    Modifiable50.BackColor = 12632256   'Replace le contrôle en gris.
End Sub

Private Sub OptionsFiltre_AfterUpdate()
    'Active ou désactive le filtre de l'option sélectionnée.
        If OptionsFiltre = 2 Then
            Me.Filter = "Ville = 'Paris'"
            Me.FilterOn = True  'Applique le filtre.
        Else
            Me.FilterOn = False 'Supprime le filtre.
        End If

End Sub

Private Sub Prenom_AfterUpdate()
'Copie la valeur du champ Prénom dans le champ Nom familier.
    If IsNull(Nomfamilier) Then
        Nomfamilier = Prenom
    End If
End Sub

Private Sub RéfTypeContact_DblClick(Cancel As Integer)
On Error GoTo Err_RéfTypeContact_DblClick
    Dim lngContactTypeID As Long

    If IsNull(Me![RéfTypeContact]) Then
        Me![RéfTypeContact].Text = ""
    Else
        lngContactTypeID = Me![RéfTypeContact]
        Me![RéfTypeContact] = Null
    End If
    DoCmd.OpenForm "Types de contacts", , , , , acDialog, "GotoNew"
    Me![RéfTypeContact].Requery
    If lngContactTypeID <> 0 Then Me![RéfTypeContact] = lngContactTypeID

Exit_RéfTypeContact_DblClick:
    Exit Sub

Err_RéfTypeContact_DblClick:
    MsgBox Err.Description
    Resume Exit_RéfTypeContact_DblClick
End Sub
Private Sub Appels_Click()
On Error GoTo Err_Appels_Click
    If IsNull(Me![RéfContact]) Then
        MsgBox "tapez les informations sur le contact avant d'appeler."
    Else
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        DoCmd.OpenForm "Appels"
    End If

Exit_Appels_Click:
    Exit Sub

Err_Appels_Click:
    MsgBox Err.Description
    Resume Exit_Appels_Click
End Sub
Private Sub Composer_Click()
On Error GoTo Err_Composer_Click

    Dim strDialStr As String
    Dim ctlPrevCtl As Control
    Const ERR_OBJNOTEXIST = 2467
    Const ERR_OBJNOTSET = 91

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

    Set ctlPrevCtl = Screen.PreviousControl

    If TypeOf ctlPrevCtl Is TextBox Then
      strDialStr = IIf(VarType(ctlPrevCtl) > V_NULL, ctlPrevCtl, "")
    ElseIf TypeOf ctlPrevCtl Is ListBox Then
      strDialStr = IIf(VarType(ctlPrevCtl) > V_NULL, ctlPrevCtl, "")
    ElseIf TypeOf ctlPrevCtl Is ComboBox Then
      strDialStr = IIf(VarType(ctlPrevCtl) > V_NULL, ctlPrevCtl, "")
    Else
      strDialStr = ""
    End If

    Application.Run "utility.wlib_AutoDial", strDialStr

Exit_Composer_Click:
    Exit Sub

Err_Composer_Click:
    If (Err = ERR_OBJNOTEXIST) Or (Err = ERR_OBJNOTSET) Then
      Resume Next
    End If
    MsgBox Err.Description
    Resume Exit_Composer_Click
End Sub
Private Sub Page1_Click()
    Me.GoToPage 1
End Sub
Private Sub Page2_Click()
    Me.GoToPage 2
End Sub



Private Sub Modifiable50_AfterUpdate()
    ' Rechercher l'enregistrement correspondant au contrôle.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[RéfContact] = " & Str(Me![Modifiable50])
    Me.Bookmark = rs.Bookmark
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Si Ctrl+1, entre la valeur Paris.
    If KeyCode = 49 And Shift = 2 Then  'Ctrl+1 a été appuyé.
        Ville = "Paris"
        CodePostal = "75000"
        Pays = "France"
    End If
End Sub
Private Function NomComplet() As String
    NomComplet = Prenom & " " & NomFamille
End Function
Public Function AdresseComplete() As String
'Retourne une chaîne contenant les nom et adresse complets.

    'S'il n'y a ni nom, ni adresse, retourne une chaîne vide.
    If IsNull(NomFamille) Or IsNull(Adresse) Then
        AdresseComplete = ""
    Else
        'Construit la chaîne d'adresse.
            AdresseComplete = NomComplet & vbNewLine & _
            Adresse & vbNewLine & _
            CodePostal & " - " & Ville
    End If
End Function
