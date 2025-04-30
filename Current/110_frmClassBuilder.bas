VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_110_frmClassBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum enuDirection
    Up
    Down
End Enum

Dim Log As New clsLog

Dim arrSelectedPackages As String



Private Sub cmdReset_Click()

    ApplyDefaultSettings

End Sub

Private Sub Form_Load()

    Log.WriteLine "Class Builder geöffnet."

    
    pagClassData.SetFocus
    Load_Packages
    ApplyDefaultSettings
    UpdateForeignKeyActivation

    
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Properties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Methods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
    

End Sub
Private Sub Form_Close()

    Log.WriteLine "Class Builder geschlossen."

End Sub
Private Sub txtClassName_AfterUpdate()

    txtClassTableName = txtClassName
    
End Sub
'---------------------------------------------------------- PACKAGES ---------------------------------------------
Private Sub Load_Packages()

      Dim intRow As Long
    Dim intCol As Long
    Dim intRows As Long
    Dim intCols As Long
    Dim varRow() As Variant
    Dim varData() As Variant

    lstPreview_Packages.RowSource = ""
    
    varData = Get_Array_FromQuery("110_qryClassBuilder_Package_SORT")

    On Error GoTo ExitSub
    intRows = UBound(varData, 1)
    intCols = UBound(varData, 2)
    ' ? 2D-Array erkannt

    For intRow = 1 To intRows
        ReDim varRow(0 To intCols - 1)
        For intCol = 1 To intCols
            varRow(intCol - 1) = varData(intRow, intCol)
        Next intCol
        lstPreview_Packages.AddItem varRow(0)
        For intCol = 1 To intCols - 1
'            lstPackages.List(lstPackages.ListCount - 1, intCol) = varRow(intCol)
        Next intCol
    Next intRow
    Exit Sub

ExitSub:
    ' Falls 1D-Array, wird hier weitergemacht
    On Error Resume Next
'    lstPackages.Clear
    For intRow = LBound(varData) To UBound(varData)
        lstPreview_Packages.AddItem varData(intRow)
    Next intRow

End Sub
Private Sub lstPreview_Packages_AfterUpdate()

    Dim Packages As Variant
    Dim lngCurrentPackage_ID As Long
    Dim strCurrentPackage_Name As String
    Dim blnCurrentPackage_Selected As Boolean

    Dim lngCounter_Packages As Long
    Dim rcsPackage_Class_Required As Recordset
    Dim rcsPackage_Properties As Recordset
    Dim rcsPackage_Methods As Recordset
    
    Packages = Access_ListBox_Get_Array(Me.Name, "lstPreview_Packages")


    If Not IsEmpty(Packages) Then
    
        For lngCounter_Packages = LBound(Packages) To UBound(Packages)
        
            strCurrentPackage_Name = Packages(lngCounter_Packages, 1)
            lngCurrentPackage_ID = DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & strCurrentPackage_Name & "'")
            blnCurrentPackage_Selected = Access_ListBox_IsValueSelected(Me.Name, "lstPreview_Packages", 0, strCurrentPackage_Name)
            
            'To-Do: Abhängigkeiten zu anderen Classen
            Set rcsPackage_Class_Required = CurrentDb.OpenRecordset( _
                "SELECT * FROM 110_tblClassBuilder_Package_Class_Required " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID)
                
            If rcsPackage_Class_Required.RecordCount > 0 Then
            
                rcsPackage_Class_Required.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", rcsPackage_Class_Required.Fields("ClassName").Value) = False Then
                            'Eintrag hinzufügen
                            lstPreviewClass_Required.AddItem rcsPackage_Class_Required.Fields("ClassName").Value
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", rcsPackage_Class_Required.Fields("ClassName").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPreviewClass_Required", rcsPackage_Class_Required.Fields("ClassName").Value
                        End If
                    
                    End If
                    
                    
                    rcsPackage_Class_Required.MoveNext
                    
                Loop While rcsPackage_Class_Required.EOF = False
            
            End If


            
            'Eigenschaften des Pakets
            Set rcsPackage_Properties = CurrentDb.OpenRecordset( _
                "SELECT * FROM 110_tblClassBuilder_Property_Draft " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID)
                
            If rcsPackage_Properties.RecordCount > 0 Then
            
                rcsPackage_Properties.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value) = False Then
                            'Eintrag hinzufügen
                            lstPreview_Properties.AddItem rcsPackage_Properties.Fields("Name").Value & ";" & _
                                DLookup("name", "110_tblClassBuilder_Property_Type", "ID = " & rcsPackage_Properties.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value
                        End If
                    
                    End If
                       
                    rcsPackage_Properties.MoveNext
                
                Loop While rcsPackage_Properties.EOF = False
                
            
            End If
            
            
            
            'Methoden des Pakets
            Set rcsPackage_Methods = CurrentDb.OpenRecordset("SELECT * FROM 110_tblClassBuilder_Method_Draft " & _
                "WHERE Package_FK = " & _
                DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & strCurrentPackage_Name & "'"))
                
            If rcsPackage_Methods.RecordCount > 0 Then
            
                rcsPackage_Methods.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Methods", rcsPackage_Methods.Fields("Name").Value) = False Then
                            lstPreview_Methods.AddItem rcsPackage_Methods.Fields("Name").Value & ";" & _
                                DLookup("name", "110_tblClassBuilder_Visability", "ID = " & rcsPackage_Methods.Fields("Visability_FK").Value) & ";" & _
                                DLookup("name", "110_tblClassBuilder_Method_Type", "ID = " & rcsPackage_Methods.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Methods", rcsPackage_Methods.Fields("Name").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPreview_Methods", rcsPackage_Methods.Fields("Name").Value
                        End If
                    
                    End If
                       
                    rcsPackage_Methods.MoveNext
                
                Loop While rcsPackage_Methods.EOF = False
                
            
            End If
        
        Next lngCounter_Packages
        
    End If
    
End Sub
'---------------------------------------------------------- REQUIRED CLASS ---------------------------------------------
Private Sub cmdRequiredClass_Add_Click()

    If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", cmbClassRequired_Add.Value) = False Then
        If cmbClassRequired_Add.Value <> "" Then
            lstPreviewClass_Required.AddItem cmbClassRequired_Add.Value
            cmbClassRequired_Add.Value = ""
            cmbClassRequired_Add.SetFocus
        Else
            MsgBox "Es wurde kein Wert eingetragen."
            End
        End If
        
    Else
    
        MsgBox "Die Klasse wurde bereits hinzugefügt."
        End
    
    End If

End Sub
Private Sub Reset_Class_Required()

    cmbClassRequired_Add = ""
    Access_ListBox_Clear Me.Name, "lstPreviewClass_Required"

End Sub
Private Sub cmdPreviewClass_Required_DeleteSelected_Click()
    Access_Listbox_SelectedItem_Delete Me.Name, "lstPreviewClass_Required"
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreviewClass_Required", cmdPreviewClass_Required_MoveUp, cmdPreviewClass_Required_MoveDown
End Sub
Private Sub cmdPreviewClass_Required_MoveDown_Click()
    Access_Listbox_SelectedItem_Move Me.Name, lstPreviewClass_Required, Down
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreviewClass_Required", cmdPreviewClass_Required_MoveUp, cmdPreviewClass_Required_MoveDown

End Sub
Private Sub cmdPreviewClass_Required_MoveUp_Click()
    Access_Listbox_SelectedItem_Move Me.Name, lstPreviewClass_Required, Up
End Sub
Private Sub lstPreviewClass_Required_Click()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreviewClass_Required", cmdPreviewClass_Required_MoveUp, cmdPreviewClass_Required_MoveDown
End Sub
Private Sub lstPreviewClass_Required_GotFocus()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreviewClass_Required", cmdPreviewClass_Required_MoveUp, cmdPreviewClass_Required_MoveDown
End Sub

'---------------------------------------------------------- PROPERTIES ---------------------------------------------
Private Sub cmdProperty_Add_Click()
    
    'To-Do: Anpassen, so, dass auch die Daten für einen Fremdschlüssel abgefragt werden
    
    If txtAddPropertyName.Value <> "" And cmbAddPropertyType.Value <> "" Then
    
        If IsNull(DLookup("ID", "110_tblClassBuilder_Property_Draft", "Name = '" & txtAddPropertyName.Value & "'")) = False Then
            MsgBox "Der Name der Property ist bereits an eine andere Property vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Property an und probieren Sie es erneut."
            Exit Sub
        End If
     
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreview_Properties", 0, txtAddPropertyName.Value) = False Then

            If chkAddProperty_IsForeignKey = True Then
            
                'To-Do: Datentyp aus der Property der Klasse, die verknüoft werden soll
                lstPreview_Properties.AddItem txtAddPropertyName.Value & ";" & _
                    cmbAddPropertyType.Column(1)
            
            
            Else

                lstPreview_Properties.AddItem txtAddPropertyName.Value & ";" & _
                    cmbAddPropertyType.Column(1)
                
            End If
            
            txtAddPropertyName.Value = ""
            cmbAddPropertyType.Value = ""
            
            txtAddPropertyName.SetFocus
            
        Else
        
            MsgBox "Eine Eigenschaft mit dem gleichen Namen wurde bereits hinzugefügt."
        
        End If
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht befüllt."
        
    End If

End Sub
Private Sub Reset_Properties()

    txtAddPropertyName = ""
    lstPreview_Properties.ColumnCount = 4
    Access_ListBox_Clear Me.Name, "lstPreview_Properties"

End Sub
Private Sub chkAddProperty_IsForeignKey_AfterUpdate()
    
    UpdateForeignKeyActivation

End Sub
Private Sub cmbAddProperty_Class_FK_AfterUpdate()

    cmbAddProperty_Property_FK.Requery

End Sub
Private Sub UpdateForeignKeyActivation()

    If chkAddProperty_IsForeignKey = True Then
    cmbAddPropertyType = ""
    cmbAddPropertyType.Enabled = False
        cmbAddProperty_Class_FK.Enabled = True
        cmbAddProperty_Property_FK.Enabled = True
    Else
        cmbAddProperty_Class_FK = ""
        cmbAddProperty_Class_FK.Enabled = False
        cmbAddProperty_Property_FK = ""
        cmbAddProperty_Property_FK.Enabled = False
        cmbAddPropertyType.Enabled = True
    End If
    
End Sub
Private Sub cmdPreviewProperties_DeleteSelected_Click()
        Access_Listbox_SelectedItem_Delete Me.Name, "lstPreview_Properties"
End Sub
Private Sub cmdPreviewProperty_MoveDown_Click()
    Access_Listbox_SelectedItem_Move Me.Name, lstPreview_Properties, Down
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Properties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
Private Sub cmdPreviewProperty_MoveUp_Click()
    Access_Listbox_SelectedItem_Move Me.Name, lstPreview_Properties, Up
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Properties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
Private Sub lstPreview_Properties_Click()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Properties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
Private Sub lstPreview_Properties_GotFocus()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Properties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
'---------------------------------------------------------- METHODS ---------------------------------------------
Private Sub cmdMethod_Add_Click()


    If txtAddPropertyName.Value <> "" And cmbAddPropertyType.Value <> "" Then
    
        If IsNull(DLookup("ID", "110_tblClassBuilder_Property_Draft", "Name = '" & txtAddPropertyName.Value & "'")) = False Then
            MsgBox "Der Name der Property ist bereits an eine andere Property vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Property an und probieren Sie es erneut."
            Exit Sub
        End If
     
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreview_Properties", 0, txtAddPropertyName.Value) = False Then

            lstPreview_Properties.AddItem txtAddPropertyName.Value & ";" & _
                cmbAddPropertyType.Column(1)
                
            txtAddPropertyName.Value = ""
            cmbAddPropertyType.Value = ""
            
            txtAddPropertyName.SetFocus
            
        Else
        
            MsgBox "Eine Eigenschaft mit dem gleichen Namen wurde bereits hinzugefügt."
        
        End If
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht befüllt."
        
    End If

End Sub
Private Sub Reset_Methods()

    txtAddMethodName = ""
    
    Access_ListBox_Clear Me.Name, "lstPreview_Methods"
    
    optMethodAddTypeSub = False
    optMethodAddTypeFunction = False
    optMethodAddPrivate = False
    optMethodAddPublic = False
    
End Sub
Private Sub optMethodAddPrivate_AfterUpdate()
    Access_OptionBox_Sync Me.Name, optMethodAddPrivate, optMethodAddPublic
End Sub
Private Sub optMethodAddPublic_AfterUpdate()
    Access_OptionBox_Sync Me.Name, optMethodAddPublic, optMethodAddPrivate
End Sub
Private Sub optMethodAddTypeFunction_AfterUpdate()
    Access_OptionBox_Sync Me.Name, optMethodAddTypeFunction, optMethodAddTypeSub
End Sub
Private Sub optMethodAddTypeSub_AfterUpdate()
    Access_OptionBox_Sync Me.Name, optMethodAddTypeSub, optMethodAddTypeFunction
End Sub
Private Sub cmdPreviewMethods_DeleteSelected_Click()
    Access_Listbox_SelectedItem_Delete Me.Name, "lstPreview_Methods"
End Sub
Private Sub cmdPreviewMethod_MoveDown_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Methods", Down
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Methods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub cmdPreviewMethod_MoveUp_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Methods", Up
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Methods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreview_Methods_Click()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Methods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreview_Methods_GotFocus()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Methods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
'---------------------------------------------------------- ENUMERATIONS ---------------------------------------------
Private Sub Reset_Enumerations()

    Access_ListBox_Clear Me.Name, "lstPreview_Enumeration"
    txtEnumerationAdd_Name = ""
    txtEnumerationAdd_Acronym = ""
End Sub
Private Sub cmdPreviewEnumerations_DeleteSelected_Click()
    Access_Listbox_SelectedItem_Delete Me.Name, "lstPreview_Enumeration"
End Sub
Private Sub cmdPreviewEnumeration_MoveDown_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Enumeration", Down
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Enumeration", cmdPreviewEnumeration_MoveUp, cmdPreviewEnumeration_MoveDown
End Sub
Private Sub cmdPreviewEnumeration_MoveUp_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Enumeration", Up
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Enumeration", cmdPreviewEnumeration_MoveUp, cmdPreviewEnumeration_MoveDown
End Sub
Private Sub lstEnumerations_Click()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Enumeration", cmdPreviewEnumeration_MoveUp, cmdPreviewEnumeration_MoveDown
End Sub
Private Sub lstEnumerations_GotFocus()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Enumeration", cmdPreviewEnumeration_MoveUp, cmdPreviewEnumeration_MoveDown
End Sub
'---------------------------------------------------------- TRANSLATION ---------------------------------------------
Private Sub Reset_Translations()

    txtEnumeration_Selected = ""
    txtEnum_Value = ""
    txtEnum_Translation = ""

    Access_ListBox_Clear Me.Name, "lstPreview_Translation"
    
End Sub

Private Sub cmdPreviewTranslations_DeleteSelected_Click()
        Access_Listbox_SelectedItem_Delete Me.Name, "lstPreview_Translation"
End Sub
Private Sub cmdPreviewTranslation_MoveDown_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Translation", Down
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Translation", cmdPreviewTranslation_MoveUp, cmdPreviewTranslation_MoveDown
End Sub
Private Sub cmdPreviewTranslation_MoveUp_Click()
    Access_Listbox_SelectedItem_Move Me.Name, "lstPreview_Translation", Up
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Translation", cmdPreviewTranslation_MoveUp, cmdPreviewTranslation_MoveDown
End Sub
Private Sub lstPreview_Translation_Click()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Translation", cmdPreviewTranslation_MoveUp, cmdPreviewTranslation_MoveDown
End Sub
Private Sub lstPreview_Translation_GotFocus()
    Access_ListBox_MovingButtons_UpdateActivation Me.Name, "lstPreview_Translation", cmdPreviewTranslation_MoveUp, cmdPreviewTranslation_MoveDown
End Sub
'---------------------------------------------------------- Class Create ---------------------------------------------
Private Sub cmdClass_Create_Click()

    Dim Class As New clsClass_Old

    Dim Properties() As Variant
    Dim Methods() As Variant
    Dim Classes() As Variant
    
    If txtClassName <> "" Then
    
        Properties = Access_ListBox_Get_Array(Me.Name, "lstPreview_Properties")
        Methods = Access_ListBox_Get_Array(Me.Name, "lstPreview_Methods")
        Classes = Access_ListBox_Get_Array(Me.Name, "lstPreviewClass_Required")
        
        Class.Build Me.txtClassName.Value, Properties(), Methods(), Classes()
    
    Else
    
        MsgBox "Es wurde kein Name für die Klasse eingetragen."
        End
    
    End If
    
End Sub

Private Sub ApplyDefaultSettings()
    
    txtClassName = ""
    
    Reset_Class_Required
    Reset_Properties
    Reset_Methods
    Reset_Enumerations
    Reset_Translations
    
End Sub




