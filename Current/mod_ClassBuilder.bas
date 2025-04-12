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

Public Sub Class_Build_New(strClassName As String, strProperties() As String)

    Dim strVBACode As String

    Log.WriteLine "Generierung der Klasse gestartet"


    strVBACode = Coding_VBA.Get_Code_Module(strClassName, True, True, strProperties(), PropertyTypes())
    Log.WriteLine "VBA Klassencode erstellt."
    

    Coding.ClassModule_CreateNew "cls" & strClassName, strVBACode
    Log.WriteLine "Klassenmodul erstellt."
     
    'Create Table

    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable("tbl_" & strClassName, strProperties(), PropertyTypes())
    Log.WriteLine "Tabelle wurde erstellt."
    
    Log.WriteLine "Klasse wurde erstellt."

End Sub