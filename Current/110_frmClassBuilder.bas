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

    Log.WriteLine "Class Builder ge�ffnet."

    
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
    
    varData = Get_Array_FromQuery("qry_Package_SORT")

    On Error GoTo ExitSub
    intRows = UBound(varData, 1)
    intCols = UBound(varData, 2)
    ' ? 2D-Array erkannt

    For intRow = 0 To intRows - 1
    
       
        lstPreview_Packages.AddItem varData(intRow + 1, 2)
        

        
    Next intRow
    
    Exit Sub

ExitSub:
    ' Falls 1D-Array, wird hier weitergemacht
    On Error Resume Next
'    lstPackages.Clear
    For intRow = LBound(varData) To UBound(varData)
        lstPreview_Packages.AddItem varData(intRow, 2)
    Next intRow

End Sub
Private Sub lstPreview_Packages_AfterUpdate()

    Dim Packages As Variant
    Dim strSQL As String
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
            lngCurrentPackage_ID = DLookup("ID", "tbl_Package", "Name = '" & strCurrentPackage_Name & "'")
            blnCurrentPackage_Selected = Access_ListBox_IsValueSelected(Me.Name, "lstPreview_Packages", 0, strCurrentPackage_Name)
            
            
            Set rcsPackage_Class_Required = CurrentDb.OpenRecordset( _
                "SELECT * FROM tbl_Class_Old WHERE ID In(" & _
                "SELECT Class_FK FROM tbl_Package_Class " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID & ")")
                
                
                
            If rcsPackage_Class_Required.RecordCount > 0 Then
            
                rcsPackage_Class_Required.MoveFirst
                
                Do
                
                    'Klassen des Pakets
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", rcsPackage_Class_Required.Fields("Name").Value) = False Then
                            'Eintrag hinzuf�gen
                            lstPreviewClass_Required.AddItem rcsPackage_Class_Required.Fields("Name").Value
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", _
                                rcsPackage_Class_Required.Fields("Name").Value) = True Then
                            'Eintrag l�schen
                            Access_ListBox_RemoveValue Me.Name, "lstPreviewClass_Required", rcsPackage_Class_Required.Fields("Name").Value
                        End If
                    
                    End If
                    
                    
                    

                    
                     
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    rcsPackage_Class_Required.MoveNext
                    
                Loop While rcsPackage_Class_Required.EOF = False
            
            End If


            
            'Eigenschaften der Klasse
            Set rcsPackage_Properties = CurrentDb.OpenRecordset("SELECT * FROM tbl_Package_Property_Draft " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID, dbOpenSnapshot)
                
            If rcsPackage_Properties.RecordCount > 0 Then
            
                rcsPackage_Properties.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value) = False Then
                            'Eintrag hinzuf�gen
                            lstPreview_Properties.AddItem rcsPackage_Properties.Fields("Name").Value & ";" & _
                                DLookup("name", "tbl_Property_Type", "ID = " & rcsPackage_Properties.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value) = True Then
                            'Eintrag l�schen
                            Access_ListBox_RemoveValue Me.Name, "lstPreview_Properties", rcsPackage_Properties.Fields("Name").Value
                        End If
                    
                    End If
                       
                    rcsPackage_Properties.MoveNext
                
                Loop While rcsPackage_Properties.EOF = False
                
            
            End If
            
            
            
            'Methoden des Pakets
            Set rcsPackage_Methods = CurrentDb.OpenRecordset("SELECT * FROM tbl_Package_Method_Draft " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID, dbOpenSnapshot)

                
            If rcsPackage_Methods.RecordCount > 0 Then
            
                rcsPackage_Methods.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Methods", rcsPackage_Methods.Fields("Name").Value) = False Then
                            lstPreview_Methods.AddItem rcsPackage_Methods.Fields("Name").Value & ";" & _
                                DLookup("name", "tbl_Visability", "ID = " & rcsPackage_Methods.Fields("Visability_FK").Value) & ";" & _
                                DLookup("name", "tbl_Method_Type", "ID = " & rcsPackage_Methods.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreview_Methods", rcsPackage_Methods.Fields("Name").Value) = True Then
                            'Eintrag l�schen
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

    If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", cmbClassRequired_Add.Column(1)) = False Then
        If cmbClassRequired_Add.Value <> "" Then
            lstPreviewClass_Required.AddItem cmbClassRequired_Add.Column(1)
            cmbClassRequired_Add.Value = ""
            cmbClassRequired_Add.SetFocus
        Else
            MsgBox "Es wurde kein Wert eingetragen."
            End
        End If
        
    Else
    
        MsgBox "Die Klasse wurde bereits hinzugef�gt."
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

    
    If (txtAddPropertyName.Value <> "" And cmbAddPropertyType.Value <> "") Or _
        (txtAddPropertyName.Value <> "" And cmbAddProperty_Class_FK.Value <> "" And cmbAddProperty_Property_FK.Value <> "") Then
    
        If IsNull(DLookup("ID", "tbl_Package_Property_Draft", "Name = '" & txtAddPropertyName.Value & "'")) = False Then
            MsgBox "Der Name der Property ist bereits an eine andere Property vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Property an und probieren Sie es erneut."
            Exit Sub
        End If
     
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreview_Properties", 0, txtAddPropertyName.Value) = False Then

            If chkAddProperty_IsForeignKey = True Then
            
                If Right(txtAddPropertyName, 3) <> "_FK" Then txtAddPropertyName = txtAddPropertyName & "_FK"
            
                lstPreview_Properties.AddItem txtAddPropertyName.Value & ";" & _
                    DLookup("DataType", "tbl_Class_Property", "Class_FK = " & cmbAddProperty_Class_FK.Column(0) & _
                        " AND Name = '" & cmbAddProperty_Property_FK.Column(1) & "'") & ";" & _
                    cmbAddProperty_Class_FK.Column(1) & ";" & _
                    cmbAddProperty_Property_FK.Column(1)
                    
                'Ben�tigter Verweis auf die Klasse
                If Access_ListBox_ContainsValue(Me.Name, "lstPreviewClass_Required", cmbAddProperty_Class_FK.Column(1)) = False Then
                    lstPreviewClass_Required.AddItem cmbAddProperty_Class_FK.Column(1)
                End If
                
                
            Else

                lstPreview_Properties.AddItem txtAddPropertyName.Value & ";" & _
                    cmbAddPropertyType.Column(1)
                
            End If
            
            txtAddPropertyName.Value = ""
            cmbAddPropertyType.Value = ""
            chkAddProperty_IsForeignKey = False
            cmbAddProperty_Property_FK = ""
            cmbAddProperty_Class_FK = ""
            
            UpdateForeignKeyActivation
            
            
            txtAddPropertyName.SetFocus
            
        Else
        
            MsgBox "Eine Eigenschaft mit dem gleichen Namen wurde bereits hinzugef�gt."
        
        End If
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht bef�llt."
        
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
    
    If cmbAddProperty_Class_FK.Value <> "" Then
    
        cmbAddProperty_Property_FK.Locked = False
        cmbAddProperty_Property_FK.Requery
        
    Else
        
        UpdateForeignKeyActivation
        
    End If

End Sub
Private Sub cmbAddProperty_Property_FK_AfterUpdate()

    If cmbAddProperty_Class_FK <> "" And cmbAddProperty_Property_FK <> "" And txtAddPropertyName = "" Then
        txtAddPropertyName = Replace(cmbAddProperty_Class_FK, "cls", "") & "_FK"
    End If

End Sub
Private Sub UpdateForeignKeyActivation()

    If chkAddProperty_IsForeignKey = True Then
    
        cmbAddPropertyType = ""
        cmbAddPropertyType.Locked = True
        cmbAddProperty_Class_FK.Locked = False
        cmbAddProperty_Property_FK.Locked = True
              
    Else
    
        cmbAddProperty_Class_FK = ""
        cmbAddProperty_Class_FK.Locked = True
        cmbAddProperty_Property_FK = ""
        cmbAddProperty_Property_FK.Locked = True
        cmbAddPropertyType.Locked = False
        
    End If
    
    'Design aktualisieren
    Access_Form_Controls_ApplyDesign Me.Name, "cmbAddProperty_Class_FK", GetDesign_AsArray_DarkMode()
    Access_Form_Controls_ApplyDesign Me.Name, "cmbAddProperty_Property_FK", GetDesign_AsArray_DarkMode()
    Access_Form_Controls_ApplyDesign Me.Name, "cmbAddPropertyType", GetDesign_AsArray_DarkMode()
    
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

    Dim strType As String
    Dim strVisabilty As String

    If txtAddMethodName.Value <> "" Then
    
        If IsNull(DLookup("ID", "tbl_Package_Method_Draft", "Name = '" & txtAddMethodName.Value & "'")) = False Then
            MsgBox "Der Name der Methode ist bereits an eine andere Methode vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Methode an und probieren Sie es erneut."
            Exit Sub
        End If
        
        'Get Visability String
        If optMethodAddPrivate = True Then
            strVisabilty = "Private"
        Else
            If optMethodAddPublic = True Then
                strVisabilty = "Public"
            End If
        End If
        
        'Get Type String
        If optMethodAddTypeFunction = True Then
            strType = "Function"
        Else
            
            If optMethodAddTypeSub = True Then
                strType = "Sub"
            End If
        End If
     
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreview_Methods", 0, txtAddMethodName.Value) = False Then

            lstPreview_Methods.AddItem txtAddMethodName.Value & ";" & _
                strType & ";" & strVisabilty
                
            txtAddMethodName.Value = ""
            optMethodAddTypeSub = False
            optMethodAddTypeFunction = False
            optMethodAddPrivate = False
            optMethodAddPublic = False
            
            txtAddPropertyName.SetFocus
            
        Else
        
            MsgBox "Eine Eigenschaft mit dem gleichen Namen wurde bereits hinzugef�gt."
        
        End If
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht bef�llt."
        
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

'---------------------------------------------------------- Class Create ---------------------------------------------
Private Sub cmdClass_Create_Click()

    Dim Class As New clsClass

    Dim Properties() As Variant
    Dim Methods() As Variant
    Dim Classes() As Variant
    
    If txtClassName <> "" Then
    
        Properties = Access_ListBox_Get_Array(Me.Name, "lstPreview_Properties")
        Methods = Access_ListBox_Get_Array(Me.Name, "lstPreview_Methods")
        Classes = Access_ListBox_Get_Array(Me.Name, "lstPreviewClass_Required")
        
        Class.Build Me.txtClassName.Value, Properties(), Methods(), Classes()
    
    Else
    
        MsgBox "Es wurde kein Name f�r die Klasse eingetragen."
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
'################################ Enumerations / Translations ##################################
'-------------------------------------- ENUMERATIONS ---------------------------------------------
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
'---------------------------------------- TRANSLATION ---------------------------------------------
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


