Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Dim Log As New clsLog
Dim Coding As New clsCoding
Dim Coding_SQL As New clsCoding_SQL
Dim Coding_VBA As New clsCoding_VBA



Public Sub Build(strClassName As String, _
    Properties() As Variant, _
    Methods() As Variant, _
    Classes() As Variant)

    Dim strVBACode As String
    
    If Access_Table_Exists("tbl_" & strClassName) = True Then
    
        MsgBox "Es existiert bereits eine Tabelle mit dem Namen." & vbNewLine & _
            "Bitte löschen Sie die Tabelle oder ändern den Namen der Klasse. "
        End
        
    End If
    
    If Access_Class_Module_Exists("cls" & strClassName) = True Then
    
        MsgBox "Es existiert bereits ein Klassenmodul mit dem Namen." & vbNewLine & _
            "Bitte löschen Sie die Tabelle oder ändern den Namen der Klasse. "
        End
        
    End If
    

    Log.WriteLine "Generierung der Klasse gestartet"
    
    
    'Create VBA Code
    strVBACode = Coding_VBA.Get_Code_ClassModule(strClassName, Properties(), Methods(), Classes())
    ClassModule_CreateNew "cls" & strClassName, strVBACode
   
     
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable("tbl_" & strClassName, Properties())
    
    
    'Erstelle die Standardformulare
    Access_Form_Create_Standard strClassName, Properties()
    
    
    Log.WriteLine "Klasse wurde erstellt."
    Log.WriteEmptyLine

End Sub

Public Sub ClassModule_CreateNew(strClassName As String, Optional strVBACode As String)
'Creates a new class module and inserts VBA Code

    Dim objVBComp As New VBComponent


    Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    objVBComp.CodeModule.AddFromString strVBACode
    objVBComp.Name = strClassName
    
    DoCmd.Save acModule, strClassName

    Log.WriteLine "Klassenmodul " & strClassName & " erstellt und VBA Code eingefügt."
    

End Sub