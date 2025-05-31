'Option Compare Database
'Option Explicit
'
'
Public Sub Access_Control_CreateFromArray( _
    strFormName As String, _
    arrControls() As Variant)

    ' Erstellt mehrere Steuerelemente automatisch untereinander auf einem Formular
    ' - strFormName: Name des Formulars
    ' - arrControls(): 2D-Array (Name, Typ)

    Dim i As Long
    Dim lngCurrentTop As Long
    Dim strControlName As String
    Dim StrControlType As String

    ' Standardwerte für Position und Größe
    Const lngStartLeft As Long = 500
    Const lngStartTop As Long = 500
    Const lngWidth As Long = 3000
    Const lngHeight As Long = 400
    Const lngSpacing As Long = 200

    On Error GoTo Fehler

    lngCurrentTop = lngStartTop

    For i = LBound(arrControls) To UBound(arrControls)

        strControlName = arrControls(i, 0)
        StrControlType = arrControls(i, 1)

        Access_Control_AddToForm _
            strFormName, _
            StrControlType, _
            strControlName, _
            lngStartLeft, _
            lngCurrentTop, _
            lngWidth, _
            lngHeight

        ' Position für das nächste Steuerelement vorbereiten
        lngCurrentTop = lngCurrentTop + lngHeight + lngSpacing

    Next i

    Exit Sub

Fehler:
    MsgBox "Fehler beim Erstellen der Steuerelemente: " & Err.Description, vbExclamation

End Sub
Public Sub Access_Control_AddToForm( _
    ByVal strFormName As String, _
    ByVal StrControlType As String, _
    ByVal strControlName As String, _
    ByVal lngLeft As Long, _
    ByVal lngTop As Long, _
    ByVal lngWidth As Long, _
    ByVal lngHeight As Long)


    Dim objForm As Access.Form
    Dim objControl As Access.Control

    On Error GoTo Fehler

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        DoCmd.OpenForm strFormName, acDesign
    End If

    Set objForm = Forms(strFormName)

    'To-Do: Überarbeiten
    ' Neues Steuerelement hinzufügen
'    Set objControl = CreateControl(strFormName, ctöTyüe, , , , lngLeft, lngTop, lngWidth, lngHeight)

    objControl.Name = strControlName
    DoCmd.Save acForm, strFormName

    Exit Sub

Fehler:
    MsgBox "Fehler beim Einfügen des Steuerelements: " & Err.Description, vbExclamation

End Sub
Public Function Translate_Properties_To_Controls(Properties() As Variant) As Variant()

    ' Wandelt ein Property-Array (Name, Typ als String) in ein Control-Array (Name, Access-ControlType) um
    ' Gibt ein 2D-Array zurück: (Name, ControlType als Integer)

    Dim lngRow As Long
    Dim varResult() As Variant
    Dim intControlType As Integer
    Dim strType As String

    ReDim varResult(LBound(Properties) To UBound(Properties), 0 To 1)

    For lngRow = LBound(Properties) To UBound(Properties)

        ' Namen übernehmen
        varResult(lngRow, 0) = Access_Form_Control_Get_Name(CStr(Properties(lngRow, 1)), _
            Translate_DataType_ToControlType(CStr(Properties(lngRow, 2))), "")
        varResult(lngRow, 1) = intControlType

    Next lngRow

    Translate_Properties_To_Controls = varResult

End Function
Private Function Translate_DataType_ToControlType(strDataType As String) As String

    Select Case strDataType

            Case "Boolean"
                Translate_DataType_ToControlType = "CheckBox"

            Case Else
                Translate_DataType_ToControlType = "TextBox" ' Standardmäßig ein Textfeld

        End Select

End Function
Public Function Access_Form_Control_Get_Name(strName As String, _
    StrControlType As String, _
    strAddition As String) As String

    Access_Form_Control_Get_Name = Access_Form_Control_Get_Prefix_ByType(StrControlType) & _
        strName & strAddition


End Function
Public Function Access_Form_Control_Get_Prefix_ByType(StrControlType As String) As String
' Gibt das passende Präfix für den angegebenen Steuerelementtyp zurück

'To-DO: Anpassen und aufräumen

    Dim strType As String

    strType = LCase(Trim(StrControlType))

    Select Case strType
        Case "TextBox", "text", "string", "number", "date", "datetime", "currency", "integer", "long", "double"
            Access_GetControlPrefixByType = "txt"
        Case "combobox", "list", "dropdown"
            Access_GetControlPrefixByType = "cbo"
        Case "listbox"
            Access_GetControlPrefixByType = "lst"
        Case "CheckBox", "boolean", "yes/no", "true/false"
            Access_GetControlPrefixByType = "chk"
        Case "commandbutton", "button"
            Access_GetControlPrefixByType = "cmd"
        Case "label"
            Access_GetControlPrefixByType = "lbl"
        Case "optiongroup"
            Access_GetControlPrefixByType = "fra"
        Case "optionbutton"
            Access_GetControlPrefixByType = "opt"
        Case "togglebutton"
            Access_GetControlPrefixByType = "tgl"
        Case "image"
            Access_GetControlPrefixByType = "img"
        Case "subform"
            Access_GetControlPrefixByType = "sfr"
        Case Else
            Access_GetControlPrefixByType = "ctl" ' Allgemeines Präfix für unbekannte Typen
    End Select

End Function
Public Sub Access_Form_Controls_ApplyDesign(strFormName As String, _
    strControlName As String, _
    varDesign As Variant)

    ' Wendet das Design aus dem varDesign-Array auf ein einzelnes Text- oder Kombinationsfeld an.
    ' Berücksichtigt Locked-Zustand und wählt entsprechende Farben.

    Dim objControl As Control

    Set objControl = Forms(strFormName).Controls(strControlName)

    Select Case objControl.ControlType

        Case acTextBox, acComboBox

            If objControl.Locked = True Then
                objControl.BackColor = varDesign(10)
                objControl.ForeColor = varDesign(11)
            Else
                objControl.BackColor = varDesign(2)
                objControl.ForeColor = varDesign(3)
            End If

            objControl.BorderColor = varDesign(4)

        Case Else
            ' Optional: andere Steuerelementtypen ignorieren oder Fehler ausgeben
            Debug.Print "Steuerelementtyp wird nicht unterstützt: " & objControl.Name

    End Select

End Sub
Public Sub Access_Form_Controls_ApplyStandardSettings(strFormName As String)

    ' Prüft alle Steuerelemente auf dem Formular und reagiert typabhängig.
    ' Bei Listboxen wird der RowSourceType auf "Value List" gesetzt.
    
    Dim objForm As Form
    Dim objControl As Control

    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    Set objForm = Forms(strFormName)

    For Each objControl In objForm.Controls
        Select Case objControl.ControlType
        
            Case acListBox
                ' Setzt ListBox auf "Value List"
                objControl.RowSourceType = "Value List"
                objControl.RowSource = ""
            
            Case acTextBox
                ' Platzhalter: TextBox-Verarbeitung folgt hier

            Case acComboBox
                ' Platzhalter: ComboBox-Verarbeitung folgt hier

            Case acCommandButton
                ' Platzhalter: Button-Verarbeitung folgt hier

            Case acLabel
                ' Platzhalter: Label-Verarbeitung folgt hier

            Case acCheckBox
                ' Platzhalter: CheckBox-Verarbeitung folgt hier

            Case acOptionGroup
                ' Platzhalter: Optionsgruppe-Verarbeitung folgt hier

            Case acTabCtl
                ' Platzhalter: TabControl-Verarbeitung folgt hier

            Case acSubform
                ' Platzhalter: Subformular-Verarbeitung folgt hier

            Case acImage
                ' Platzhalter: Bild-Steuerelement-Verarbeitung folgt hier

            Case Else
                ' Unbekannter oder nicht behandelter Steuerelementtyp
                
        End Select
    Next objControl

    DoCmd.Save
    DoCmd.Close acForm, strFormName, acSaveNo

End Sub