VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Package_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Log As New clsLog
Dim Package As New clsPackage
Dim Class As New clsClass
Dim Package_Class As New clsPackage_Class


Private Sub cmbPackageManage_Name_AfterUpdate()

    Form_Update

End Sub

Private Sub cmdClass_AddExisting_Click()

    If Access_ListBox_ContainsValue(Me.Name, _
        "lstConnectedProperties", _
        cmbClass_AddExisting) = False Then
        
            Package_Class.Name = cmbClass_AddExisting.Column(1)
            Package_Class.Package_FK = cmbPackageManage_Name
            Package_Class.Class_FK = cmbClass_AddExisting
            Package_Class.Save
        
    End If
    
    Log.WriteLine "Die Klasse '" & cmbClass_AddExisting.Column(1) & _
        "' wurde dem Paket '" & cmbPackageManage_Name.Column(1) & "' zugeordnet."
    
    cmbClass_AddExisting = ""
    
    Form_Update
    
    
End Sub

Private Sub cmdClass_Create_Click()

    DoCmd.OpenForm "110_frmClassBuilder", acNormal

    If txtClass_Create_Name.Value <> "" Then
           Forms("110_frmClassBuilder").Controls("txtClassName").Value = txtClass_Create_Name
    End If

End Sub

Private Sub cmdClass_DeleteSelected_Click()
'WORKING ON

    Dim varData As Variant
    Dim lngIndex As Long
    
    varData = Access_ListBox_Get_Array_Selected(Me.Name, "lstConnectedClasses")
    
    For lngIndex = LBound(varData) To UBound(varData)
        
        If IsNull(DLookup("ID", "tbl_Package_Class", "Name = '" & varData(lngIndex) & _
            "' AND Package_FK = " & cmbPackageManage_Name)) = False Then
            
            With Package_Class
                .LoadByID DLookup("ID", "tbl_Package_Class", "Name = '" & varData(lngIndex) & _
                        "' AND Package_FK = " & cmbPackageManage_Name)
                .Delete
            End With
        
        End If
        
    Next lngIndex
    
    Form_Update

End Sub

Private Sub cmdPackage_Add_Click()

    Dim Package As New clsPackage
    
    Log.WriteLine "Paket " & txtPackageAdd_Name & " wird erstellt."
    
    Package.Name = txtPackageAdd_Name
    Package.Save
    
    Log.WriteLine "Paket " & txtPackageAdd_Name & " wurde erstellt."
    

    Form_Clear
    

End Sub

Private Sub Form_Clear()

    txtPackageAdd_Name = ""
    cmbPackageManage_Name = ""
    Form_Update
    

End Sub
Private Sub Form_Update()

    Dim varData As Variant

    If cmbPackageManage_Name.Value <> "" Then
    
        'Klassen
            Access_ListBox_Fill_FromArray Me.Name, _
                "lstConnectedClasses", _
                Array_GetFromSQL( _
                    "SELECT Name FROM tbl_Package_Class " & _
                    "WHERE Package_FK = " & cmbPackageManage_Name)

        'Standard Properties
        Access_ListBox_Fill_FromArray Me.Name, _
            "lstConnectedProperties", _
            Array_GetFromSQL( _
                "SELECT Name FROM tbl_Package_Property_Draft " & _
                "WHERE Package_FK = " & cmbPackageManage_Name)
                
        'Standard Methods
        Access_ListBox_Fill_FromArray Me.Name, _
            "lstConnectedProperties", _
            Array_GetFromSQL( _
                "SELECT Name FROM tbl_Package_Method_Draft " & _
                "WHERE Package_FK = " & cmbPackageManage_Name)
        
    End If
    
    Requery
    Recalc

End Sub

Private Sub Form_Close()

    Log.WriteLine "Paketmanager geschlossen."

End Sub

Private Sub Form_Load()

    Log.WriteLine "Paketmanager geöffnet."

End Sub
