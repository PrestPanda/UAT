Option Compare Database
Option Explicit

Dim Log As New clsLog
Dim Coding As New clsCoding
Dim Coding_SQL As New clsCoding_SQL
Dim Coding_VBA As New clsCoding_VBA
Public Sub Class_Build(strClassName As String, strProperties() As String, PropertyTypes() As enuCoding_Variable_Types)

    Dim strVBACode As String

    Log.WriteLine "Generierung der Klasse gestartet"

    'Create Module

    strVBACode = Coding_VBA.Get_Code_Module(strClassName, True, True, strProperties(), PropertyTypes())
    Log.WriteLine "VBA Klassencode erstellt."
    

    Coding.ClassModule_CreateNew "cls" & strClassName, strVBACode
    Log.WriteLine "Klassenmodul erstellt."
     
    'Create Table

    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable("tbl_" & strClassName, strProperties(), PropertyTypes())
    Log.WriteLine "Tabelle wurde erstellt."
    
    Log.WriteLine "Klasse wurde erstellt."

End Sub

Sub ClassBuild_Class_Property()

    Dim strProperties(5) As String
    Dim PropertyTypes(5) As enuCoding_Variable_Types
    Dim strClassName As String

    
    

    strClassName = "Class_Property"

    strProperties(0) = "ID"
    strProperties(1) = "Name"
    strProperties(2) = "IsAddedInCode"
    strProperties(3) = "Code"
    strProperties(4) = "Class_FK"
    strProperties(5) = "Active"



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtString
    PropertyTypes(2) = vtBoolean
    PropertyTypes(3) = vtString
    PropertyTypes(4) = vtlong
    PropertyTypes(5) = vtBoolean
    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

End Sub
Sub ClassBuild_Component()


   
    Dim strProperties(8) As String
    Dim PropertyTypes(8) As enuCoding_Variable_Types
    Dim strClassName As String


    strClassName = "Component"

    strProperties(0) = "ID"
    strProperties(1) = "Name"
    strProperties(2) = "Description"
    strProperties(3) = "Type_FK"
    strProperties(4) = "Application_Part"
    strProperties(5) = "CreateTS"
    strProperties(6) = "LastScanTS"
    strProperties(7) = "LastUpdateTS"
    strProperties(8) = "Active"


    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtString
    PropertyTypes(2) = vtStringLong
    PropertyTypes(3) = vtlong
    PropertyTypes(4) = vtlong
    PropertyTypes(5) = vtDate
    PropertyTypes(6) = vtDate
    PropertyTypes(7) = vtDate
    PropertyTypes(8) = vtBoolean
    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

    

End Sub
Sub ClassBuild_Component_Table()


   
    Dim strProperties(6) As String
    Dim PropertyTypes(6) As enuCoding_Variable_Types
    Dim strClassName As String


    strClassName = "Component_Table"

    strProperties(0) = "ID"
    strProperties(1) = "ColumnCount"
    strProperties(2) = "SizeKB"
    strProperties(3) = "IsSystemTable"
    strProperties(4) = "IsLinkedTable"
    strProperties(5) = "LinkedTable_Path"
    strProperties(6) = "Component_FK"



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtlong
    PropertyTypes(2) = vtlong
    PropertyTypes(3) = vtBoolean
    PropertyTypes(4) = vtBoolean
    PropertyTypes(5) = vtString
    PropertyTypes(6) = vtlong
    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

    

End Sub
Sub ClassBuild_Component_Query()


   
    Dim strProperties(4) As String
    Dim PropertyTypes(4) As enuCoding_Variable_Types
    Dim strClassName As String


    strClassName = "Component_Query"

    strProperties(0) = "ID"
    strProperties(1) = "SQL"
    strProperties(2) = "IsUnion"
    strProperties(3) = "Type_FK"
    strProperties(4) = "Component_FK"



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtStringLong
    PropertyTypes(2) = vtBoolean
    PropertyTypes(3) = vtlong
    PropertyTypes(4) = vtlong

    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

    

End Sub
Sub ClassBuild_Component_Form()


   
    Dim strProperties(9) As String
    Dim PropertyTypes(9) As enuCoding_Variable_Types
    Dim strClassName As String


    strClassName = "Component_Form"

    strProperties(0) = "ID"
    strProperties(1) = "Code"
    strProperties(2) = "LOC"
    strProperties(3) = "Control_Count"
    strProperties(4) = "Setting_AutoCenter"
    strProperties(5) = "Setting_NavigationPanes"
    strProperties(6) = "Setting_PopUp"
    strProperties(7) = "Setting_ScrollBars"
    strProperties(8) = "Setting_StandardView"
    strProperties(9) = "Component_FK"
    



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtStringLong
    PropertyTypes(2) = vtlong
    PropertyTypes(3) = vtlong
    PropertyTypes(4) = vtBoolean
    PropertyTypes(5) = vtBoolean
    PropertyTypes(6) = vtBoolean
    PropertyTypes(7) = vtBoolean
    PropertyTypes(8) = vtString
    PropertyTypes(9) = vtlong

    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

    

End Sub
Sub ClassBuild_Component_Module()


   
    Dim strProperties(5) As String
    Dim PropertyTypes(5) As enuCoding_Variable_Types
    Dim strClassName As String


    strClassName = "Component_Module"

    strProperties(0) = "ID"
    strProperties(1) = "Code"
    strProperties(2) = "LOC"
    strProperties(3) = "IsClassModule"
    strProperties(4) = "Instancing"
    strProperties(5) = "Component_FK"
    



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtStringLong
    PropertyTypes(2) = vtlong
    PropertyTypes(3) = vtBoolean
    PropertyTypes(4) = vtString
    PropertyTypes(5) = vtlong


    
    Class_Build strClassName, strProperties(), PropertyTypes()
    

    

End Sub
Sub ClassBuild_Application()


   
    Dim strProperties(3) As String
    Dim PropertyTypes(3) As enuCoding_Variable_Types
    Dim strClassName As String
    Dim strVBACode As String
    Dim strTableName As String
    
    

    strClassName = "Application"
    strTableName = "tbl_" & strClassName
    

    strProperties(0) = "ID"
    strProperties(1) = "Name"
    strProperties(2) = "Description"
    strProperties(3) = "Active"

 



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtString
    PropertyTypes(2) = vtStringLong
    PropertyTypes(3) = vtBoolean


    
    Log.WriteLine "Generierung der Klasse gestartet"

    'Create Module
    strVBACode = Coding_VBA.Get_Code_Module(strClassName, True, True, strProperties(), PropertyTypes())
    
    Coding.ClassModule_CreateNew "cls" & strClassName, strVBACode
                                    
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable(strTableName, strProperties(), PropertyTypes())

    

End Sub
Sub ClassBuild_Application_Part()


   
    Dim strProperties(17) As String
    Dim PropertyTypes(17) As enuCoding_Variable_Types
    Dim strClassName As String
    Dim strVBACode As String
    Dim strTableName As String
    
    

    strClassName = "Application_Part"
    strTableName = "tbl_" & strClassName
    

    strProperties(0) = "ID"
    strProperties(1) = "Name"
    strProperties(2) = "Description"
    strProperties(3) = "Path"
    strProperties(4) = "Type_Programm_FK"
    strProperties(5) = "Type_Functional_FK"
    strProperties(6) = "Application_FK"
    strProperties(7) = "LOC"
    strProperties(8) = "Table_Count"
    strProperties(9) = "Query_Count"
    strProperties(10) = "Form_Count"
    strProperties(11) = "Form_LOC"
    strProperties(12) = "Module_Count"
    strProperties(13) = "Module_LOC"
    strProperties(14) = "ClassModule_Count"
    strProperties(15) = "ClassModule_LOC"
    strProperties(16) = "Active"
    strProperties(17) = "LastScanTS"
 



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtString
    PropertyTypes(2) = vtStringLong
    PropertyTypes(3) = vtString
    PropertyTypes(4) = vtlong
    PropertyTypes(5) = vtlong
    PropertyTypes(6) = vtlong
    PropertyTypes(7) = vtlong
    PropertyTypes(8) = vtlong
    PropertyTypes(9) = vtlong
    PropertyTypes(10) = vtlong
    PropertyTypes(11) = vtlong
    PropertyTypes(12) = vtlong
    PropertyTypes(13) = vtlong
    PropertyTypes(14) = vtlong
    PropertyTypes(15) = vtlong
    PropertyTypes(16) = vtBoolean
    PropertyTypes(17) = vtDate

    
    Log.WriteLine "Generierung der Klasse """ & strClassName & """ gestartet"

    'Create Module
    strVBACode = Coding_VBA.Get_Code_Module(strClassName, True, True, strProperties(), PropertyTypes())
    
    Coding.ClassModule_CreateNew "cls" & strClassName, strVBACode
                                    
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable(strTableName, strProperties(), PropertyTypes())

    

End Sub