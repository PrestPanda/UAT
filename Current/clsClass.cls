Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Dim Log As New clsLog
Dim Coding As New clsCoding
Dim Coding_SQL As New clsCoding_SQL
Dim Coding_VBA As New clsCoding_VBA



Public Sub Build(strClassName As String, Properties() As Variant, Methods() As Variant)

    Dim strVBACode As String
    

    Log.WriteLine "Generierung der Klasse gestartet"
    
    
    'Create VBA Code
    strVBACode = Coding_VBA.Get_Code_ClassModule(strClassName, Properties(), Methods())
    ClassModule_CreateNew strClassName, strVBACode
   
     
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable("tbl_" & strClassName, Properties())
    
    
    Log.WriteLine "Klasse wurde erstellt."

End Sub

Public Sub ClassModule_CreateNew(strClassName As String, Optional strVBACode As String)
'Creates a new class module and inserts VBA Code

    Dim objVBComp As New VBComponent


    Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    objVBComp.CodeModule.AddFromString strVBACode
    objVBComp.Name = strClassName

    Log.WriteLine "Klassenmodul " & strClassName & " erstellt und VBA Code eingefügt."
    

End Sub