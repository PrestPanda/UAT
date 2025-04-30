Option Compare Database
Option Explicit

Public Sub Access_OptionBox_Sync( _
    strFormName As String, _
    objChangedOption As Object, _
    objOtherOption As Object)

' Wenn das geänderte Optionsfeld aktiviert wurde,
' wird das andere automatisch deaktiviert

    
    Dim frm As Access.Form
    Set frm = Forms(strFormName)

    If frm.Controls(objChangedOption.Name).Value = True Then
        frm.Controls(objOtherOption.Name).Value = False
    Else
        frm.Controls(objOtherOption.Name).Value = True
    End If

End Sub