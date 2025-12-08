Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
    If Me.OpenArgs = "GotoNew" And Not IsNull(Me![RéfTypeContact]) Then
        DoCmd.DoMenuItem acFormBar, 3, 0, , acMenuVer70
    End If
End Sub
