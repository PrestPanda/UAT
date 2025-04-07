Option Compare Database
Option Explicit

Dim Writing As New clsWriting
Dim Coding As New clsCoding
Dim Coding_VBA As New clsCoding_VBA
Dim Coding_SQL As New clsCoding_SQL
Dim Log As New clsLog
Sub Test_ClassBuild()


   
    Dim strProperties(7) As String
    Dim PropertyTypes(7) As enuCoding_Variable_Types
    Dim strClassName As String
    Dim strVBACode As String
    Dim strTableName As String
    
    

    strClassName = "Testing"
    strTableName = "tbl_" & strClassName
    

    strProperties(0) = "ID"
    strProperties(1) = "Name"
    strProperties(2) = "Description"
    strProperties(3) = "Path"
    strProperties(4) = "Type_Programm_FK"
    strProperties(5) = "Type_Functional_FK"
    strProperties(6) = "Application_FK"
    strProperties(7) = "Active"
 



    PropertyTypes(0) = vtAuto
    PropertyTypes(1) = vtString
    PropertyTypes(2) = vtString
    PropertyTypes(3) = vtString
    PropertyTypes(4) = vtlong
    PropertyTypes(5) = vtlong
    PropertyTypes(6) = vtlong
    PropertyTypes(7) = vtBoolean

    
    Log.WriteLine "Generierung der Klasse gestartet"

    'Create Module
    strVBACode = Coding_VBA.Get_Code_Module(strClassName, True, True, strProperties(), PropertyTypes())
    
    Coding.ClassModule_CreateNew "cls" & strClassName, strVBACode
                                    
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable(strTableName, strProperties(), PropertyTypes())

    

End Sub

Sub Test_Writing_ProgressBar()


    
    Dim lngItems As Long
    Dim lngItemsTotal As Long
    
    lngItems = 19
    lngItemsTotal = 87
    
    Log.WriteLine Writing.Get_ProgressBar(lngItems, lngItemsTotal, pboDebug, pbdesDot)
    Log.WriteLine Writing.Get_ProgressBar(lngItems, lngItemsTotal, pboDebug, pbdesHash)
    Log.WriteLine Writing.Get_ProgressBar(lngItems, lngItemsTotal, pboDebug, pbdesPointer)


End Sub
Sub Test_Writing_Separator()

    
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesDash)
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesEqual)
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesStars)
    
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesCustom, "X")
    
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesDash)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesEqual)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesStars)
    
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesDash, , True)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesEqual, , True)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesStars, , True)

    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesCustom, "C", True)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesCustom, "=", True)
    
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesDash)
    Log.WriteLine Writing.Get_Separator_WithText("Überschrift", pboDebug, sepdesDash, , True)
    Log.WriteLine Writing.Get_Separator(pboDebug, sepdesDash)
    
End Sub
Sub Test_Log()

    Dim Log As New clsLog
    
    Log.WriteLine "Logger initialisiert"
    Log.WriteLine Writing.Get_Message_Welcome_User("Jan")
    
    Test_Table_Print_Information
    Test_Query_Print_Information
    Test_Form_Print_Information
    Test_Module_Print_Information

End Sub
Private Sub Test_Table_Print_Information()


    Dim cTable As New clsComponent_Table
    Dim Log As New clsLog
    
    cTable.LoadTablePropertiesFromTable "tbl_Class"
    Log.WriteLine (cTable.Get_Table_Header)
    Log.WriteLine (cTable.Get_Table_Columns)

End Sub
Private Sub Test_LogTables()

    Dim cTable As New clsComponent_Table

    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    
    Set db = CurrentDb
    
    ' Iteriere durch jede Tabelle in der Datenbank
    For Each tbl In db.TableDefs
        ' Prüfe, ob es sich nicht um eine Systemtabelle handelt
        If Left(tbl.Name, 4) <> "MSys" Then
            cTable.LoadTablePropertiesFromTable tbl.Name
            If cTable.DB_Check = False Then
                cTable.DB_Insert
            End If
        End If
    Next tbl
    
    ' Speicher freigeben
    Set tbl = Nothing
    Set db = Nothing


End Sub

Private Sub Test_Query_Print_Information()

    Dim cQuery As New clsComponent_Query
    Dim Log As New clsLog
    
    cQuery.LoadQueryProperties ("qry_ClassBuilder_Properties_CurrentUser")
    Log.WriteLine (cQuery.Get_Query_Header)
    Log.WriteLine (cQuery.Get_Query_Fields)


End Sub
Private Sub Test_Form_Print_Information()


    Dim cForm As New clsComponent_Form
    Dim Log As New clsLog
    
    cForm.LoadFormProperties ("frm_Log")
    Log.WriteLine (cForm.Get_Form_Header)
    Log.WriteLine (cForm.Get_Form_Controls)

End Sub
Private Sub Test_Module_Print_Information()


    Dim cModule As New clsComponent_Module
    Dim Log As New clsLog
    
    cModule.LoadModuleProperties ("mod_Testing")
    Log.WriteLine (cModule.Get_Module_Header)
    Log.WriteLine (cModule.Get_Module_Procedures)

End Sub