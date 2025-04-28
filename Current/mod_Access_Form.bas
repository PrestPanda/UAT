Option Compare Database
Option Explicit

Public Sub Access_Form_Create_Standard(strClassName As String)

    Access_Form_Create_Add strClassName

End Sub
Private Sub Access_Form_Create_Add(strClassName As String)

    Dim strFormName As String
    
    strFormName = "frm" & strClassName & "_Add"

    Access_Form_CreateNew strFormName
    Access_Form_Settings_PopUp strFormName

End Sub

Private Sub Access_Form_CreateNew(strFormName As String)

    Dim objFormNew As Object
    Dim strOldName As String
    


    ' Neues Formular erstellen
    Set objFormNew = CreateForm

    ' Erstmal speichern unter dem TEMP-Namen
    DoCmd.Save acForm, objFormNew.Name
    strOldName = objFormNew.Name
    ' Formular schlieﬂen
    DoCmd.Close acForm, objFormNew.Name

    ' Danach erst umbenennen
    Access_RenameForm strOldName, strFormName

    Exit Sub

Fehler:
    MsgBox "Fehler beim Erstellen oder Umbenennen des Formulars: " & Err.Description, vbExclamation


End Sub
Public Sub Access_Form_Settings_PopUp(strFormName As String)

    Dim objForm As Access.Form

    On Error GoTo Fehler

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        DoCmd.OpenForm strFormName, acDesign
    End If

    Set objForm = Forms(strFormName)

    With objForm
        .RecordSelectors = False
        .PopUp = True
        .NavigationButtons = False
    End With

    DoCmd.Save acForm, strFormName
    DoCmd.Close acForm, strFormName

    Exit Sub

Fehler:
    MsgBox "Fehler beim Setzen der Formulareinstellungen: " & Err.Description, vbExclamation

End Sub
Public Sub Access_RenameForm( _
    strOldFormName As String, _
    strNewFormName As String)

    ' Benennt ein vorhandenes Formular um

    On Error GoTo Fehler

    ' Formular muss geschlossen sein, um es umzubenennen
    If CurrentProject.AllForms(strOldFormName).IsLoaded Then
        DoCmd.Close acForm, strOldFormName
    End If

    DoCmd.Rename strNewFormName, acForm, strOldFormName

    Exit Sub

Fehler:
    MsgBox "Fehler beim Umbenennen des Formulars: " & Err.Description, vbExclamation

End Sub