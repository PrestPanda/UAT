Option Compare Database
Option Explicit

Public Sub Access_Control_AddToForm( _
    strFormName As String, _
    ctlType As AcControlType, _
    strControlName As String, _
    Optional strLabelCaption As String = "")

    ' Fügt ein neues Steuerelement auf einem Formular hinzu
    ' - ctlType: z. B. acTextBox, acComboBox, acListBox
    ' - strControlName: Name des neuen Steuerelements
    ' - strLabelCaption: (Optional) Text für das zugehörige Beschriftungsfeld

    Dim objForm As Access.Form
    Dim objControl As Access.Control

    On Error GoTo Fehler

    ' Formular im Entwurfsmodus öffnen, wenn nicht bereits offen
    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        DoCmd.OpenForm strFormName, acDesign
    End If

    Set objForm = Forms(strFormName)

    ' Neues Steuerelement hinzufügen
    Set objControl = CreateControl( _
                        strFormName, _
                        ctlType, _
                        , , , 100, 100, 2000, 400)

    ' Eigenschaften setzen
    With objControl
        .Name = strControlName
    End With

    ' Optional Beschriftung anpassen (falls vorhanden)
    If strLabelCaption <> "" Then
        objControl.Controls(0).Caption = strLabelCaption
    End If

    ' Formular speichern
    DoCmd.Save acForm, strFormName

    Exit Sub

Fehler:
    MsgBox "Fehler beim Einfügen des Steuerelements: " & Err.Description, vbExclamation

End Sub