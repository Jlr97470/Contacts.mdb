Attribute VB_Name = "Code global"
Option Compare Database

Function IsLoaded(ByVal strFormName As String) As Integer
 ' Returns True if the specified form is open in Form view or Datasheet view.

    Const conObjStateClosed = 0
    Const conDesignView = 0

    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If

End Function
