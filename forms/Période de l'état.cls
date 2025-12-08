Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Open(Cancel As Integer)
    Me.Caption = Me.OpenArgs
    Me![Date premier appel] = Date - Weekday(Date, 2) + 1
    Me![Date dernier appel] = Me![Date premier appel] + 6
End Sub
Private Sub Aperçu_Click()
    If IsNull([Date premier appel]) Or IsNull([Date dernier appel]) Then
        MsgBox "Vous devez taper une date de début et une date de fin."
        DoCmd.GoToControl "Date premier appel"
    Else
        If [Date premier appel] > [Date dernier appel] Then
            MsgBox "La date de fin doit être postérieure à la date de début."
            DoCmd.GoToControl "Date premier appel"
        Else
            Me.Visible = False
        End If
    End If
End Sub
