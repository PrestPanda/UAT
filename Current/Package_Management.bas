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


Private Sub cmbPackageManage_Name_AfterUpdate()

    Form_Update

End Sub

Private Sub cmdClass_Create_Click()

    DoCmd.OpenForm "110_frmClassBuilder", acNormal

    If txtClass_Create_Name.Value <> "" Then
           Forms("110_frmClassBuilder").Controls("txtClassName").Value = txtClass_Create_Name
    End If

End Sub

Private Sub cmdPackage_Add_Click()

    Dim Package As New clsPackage
    
    Log.WriteLine "Klasse " & txtPackageAdd_Name & " wird erstellt."
    
    Package.Name = txtPackageAdd_Name
    Package.Save
    
    Log.WriteLine "Klasse " & txtPackageAdd_Name & " wurde erstellt."
    

    Form_Clear
    

End Sub

Private Sub Form_Clear()

    txtPackageAdd_Name = ""
    cmbPackageManage_Name = ""
    Form_Update
    

End Sub
Private Sub Form_Update()

    If cmbPackageManage_Name.Value <> "" Then

        Access_ListBox_Fill_FromArray Me.Name, _
            "lstConnectedClasses", _
            Array_GetFromSQL( _
                "SELECT ID, Name FROM tbl_Package_Class " & _
                "WHERE Package_FK = " & cmbPackageManage_Name.Value)
        
    End If
    
    Requery
    Recalc

End Sub
