Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim Log As New clsLog
Dim Coding As New clsCoding
Dim Coding_SQL As New clsCoding_SQL
Dim Coding_VBA_Write As New clsCoding_VBA_Write



Private lngID As Long
Private strName As String
Private datCreateTS As Date
Private lngCreateUser_FK As Long
Private datLastUpdateTS As Date
Private lngLastUpdateUser_FK As Long
Private blnIsActive As Boolean
Private datLastActiveTS As Date
Private lngLastActiveUser_FK As Long
Private lngLOC As Long
Private lngProperty_Count As Long
Private lngMethod_Count As Long

Public Property Get ID() As Long

    ID = lngID

End Property
Public Property Let ID(lngID_1 As Long)

    lngID = lngID_1

End Property
Public Property Get Name() As String

    Name = strName

End Property
Public Property Let Name(strName_1 As String)

    strName = strName_1

End Property
Public Property Get CreateTS() As Date

    CreateTS = datCreateTS

End Property
Public Property Let CreateTS(datCreateTS_1 As Date)

    datCreateTS = datCreateTS_1

End Property
Public Property Get CreateUser_FK() As Long

    CreateUser_FK = lngCreateUser_FK

End Property
Public Property Let CreateUser_FK(lngCreateUser_FK_1 As Long)

    lngCreateUser_FK = lngCreateUser_FK_1

End Property
Public Property Get LastUpdateTS() As Date

    LastUpdateTS = datLastUpdateTS

End Property
Public Property Let LastUpdateTS(datLastUpdateTS_1 As Date)

    datLastUpdateTS = datLastUpdateTS_1

End Property
Public Property Get LastUpdateUser_FK() As Long

    LastUpdateUser_FK = lngLastUpdateUser_FK

End Property
Public Property Let LastUpdateUser_FK(lngLastUpdateUser_FK_1 As Long)

    lngLastUpdateUser_FK = lngLastUpdateUser_FK_1

End Property
Public Property Get IsActive() As Boolean

    IsActive = blnIsActive

End Property
Public Property Let IsActive(blnIsActive_1 As Boolean)

    blnIsActive = blnIsActive_1

End Property
Public Property Get LastActiveTS() As Date

    LastActiveTS = datLastActiveTS

End Property
Public Property Let LastActiveTS(datLastActiveTS_1 As Date)

    datLastActiveTS = datLastActiveTS_1

End Property
Public Property Get LastActiveUser_FK() As Long

    LastActiveUser_FK = lngLastActiveUser_FK

End Property
Public Property Let LastActiveUser_FK(lngLastActiveUser_FK_1 As Long)

    lngLastActiveUser_FK = lngLastActiveUser_FK_1

End Property
Public Property Get LOC() As Long

    LOC = lngLOC

End Property
Public Property Let LOC(lngLOC_1 As Long)

    lngLOC = lngLOC_1

End Property
Public Property Get Property_Count() As Long

    Property_Count = lngProperty_Count

End Property
Public Property Let Property_Count(lngProperty_Count_1 As Long)

    lngProperty_Count = lngProperty_Count_1

End Property
Public Property Get Method_Count() As Long

    Method_Count = lngMethod_Count

End Property
Public Property Let Method_Count(lngMethod_Count_1 As Long)

    lngMethod_Count = lngMethod_Count_1

End Property

Public Sub Reset()

    lngID = 0
    strName = ""
    datCreateTS = 0
    lngCreateUser_FK = 0
    datLastUpdateTS = 0
    lngLastUpdateUser_FK = 0
    blnIsActive = False
    datLastActiveTS = 0
    lngLastActiveUser_FK = 0
    lngLOC = 0
    lngProperty_Count = 0
    lngMethod_Count = 0

End Sub

Public Sub Load()

    DB_Load

End Sub
Public Sub Save()

    If Exists = True Then
        DB_Update
    Else
        DB_Insert
    End If

End Sub
Public Function Exists() As Boolean

    If lngID = 0 Then
        Exists = False
    Else
        If DB_Check = True Then
            Exists = True
        Else
            Exists = False
        End If
    End If

End Function
Public Sub Delete()

    If Exists = True Then
        DB_Delete
    End If

End Sub
Public Sub LoadByID(lngID_2 As Long)

    lngID = lngID_2
    Load
    

End Sub
Private Sub DB_Activate()

    If DB_Check = True Then
    
        blnIsActive = True
        DB_Update
    
    End If

End Sub
Private Function DB_Check() As Boolean

    If lngID = 0 Then
    
        DB_Check = False
    
    Else
    
        If IsNull(DLookup("ID", "tbl_Class", "ID =" & lngID)) = False Then
            DB_Check = True
        Else
            DB_Check = False
        End If
    
    End If

End Function
Private Sub DB_Deactivate()

    If DB_Check = True Then
    
        blnIsActive = False
        DB_Update
    
    End If

End Sub
Private Sub DB_Delete()

    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Class WHERE ID =" & lngID
    
    DoCmd.SetWarnings True
    
    Reset

End Sub
Private Sub DB_Insert()

    If IsNull(DMax("ID", "tbl_Class")) = False Then
        lngID = DMax("ID", "tbl_Class") + 1
    End If
    
    datCreateTS = Now()
    datLastUpdateTS = Now()
    blnIsActive = True
    
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Class " & _
        "(ID, Name, CreateTS, CreateUser_FK, LastUpdateTS, LastUpdateUser_FK, IsActive, LastActiveTS, LastActiveUser_FK, LOC, Property_Count, Method_Count) " & _
        "VALUES (" & lngID & ", '" & strName & "', " & Coding_SQL.Get_CDate_ByDate_AsString(datCreateTS) & ", " & lngCreateUser_FK & ", " & Coding_SQL.Get_CDate_ByDate_AsString(datLastUpdateTS) & ", " & lngLastUpdateUser_FK & ", " & CInt(blnIsActive) & ", " & Coding_SQL.Get_CDate_ByDate_AsString(datLastActiveTS) & ", " & lngLastActiveUser_FK & ", " & lngLOC & ", " & lngProperty_Count & ", " & lngMethod_Count & ")"
    
    DoCmd.SetWarnings True
    
    lngID = DMax("ID", "tbl_Class")
    

End Sub
Private Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    If DB_Check = True Then
    
        strSQL = "SELECT * FROM tbl_Class WHERE ID = " & lngID
    
        Reset
    
        Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    
        If Not rs.EOF Then
            lngID = rs![ID]
            strName = rs![Name]
            datCreateTS = rs![CreateTS]
            lngCreateUser_FK = rs![CreateUser_FK]
            datLastUpdateTS = rs![LastUpdateTS]
            lngLastUpdateUser_FK = rs![LastUpdateUser_FK]
            blnIsActive = rs![IsActive]
            datLastActiveTS = rs![LastActiveTS]
            lngLastActiveUser_FK = rs![LastActiveUser_FK]
            lngLOC = rs![LOC]
            lngProperty_Count = rs![Property_Count]
            lngMethod_Count = rs![Method_Count]
        End If
    
        rs.Close
        Set rs = Nothing
    
    End If
    

End Sub
Private Sub DB_Update()

    datLastUpdateTS = Now()
    
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Class SET Name = '" & strName & "', CreateTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datCreateTS) & ", CreateUser_FK = " & lngCreateUser_FK & ", LastUpdateTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datLastUpdateTS) & ", LastUpdateUser_FK = " & lngLastUpdateUser_FK & ", IsActive = " & CInt(blnIsActive) & ", LastActiveTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datLastActiveTS) & ", LastActiveUser_FK = " & lngLastActiveUser_FK & ", LOC = " & lngLOC & ", Property_Count = " & lngProperty_Count & ", Method_Count = " & lngMethod_Count & " WHERE ID = " & lngID
    
    DoCmd.SetWarnings True

End Sub
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
    strVBACode = Coding_VBA_Write.Get_Code_ClassModule(strClassName, Properties(), Methods(), Classes())
    ClassModule_CreateNew "cls" & strClassName, strVBACode
   
    'Create Table
    DoCmd.RunSQL Coding_SQL.Get_DB_CreateTable("tbl_" & strClassName, Properties())
    
    'Create Standardforms
    'Access_Form_Create_Standard strClassName, Properties()
    
    
    Log.WriteLine "Klasse wurde erstellt."
    Log.WriteEmptyLine

End Sub
Public Sub ClassModule_CreateNew(strClassName As String, _
    Optional strVBACode As String)
'Creates a new class module and inserts VBA Code

    Dim objVBComp As New VBComponent


    Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    objVBComp.CodeModule.AddFromString strVBACode
    objVBComp.Name = strClassName
    
    DoCmd.Save acModule, strClassName

    Log.WriteLine "Klassenmodul " & strClassName & " erstellt und VBA Code eingefügt."
    

End Sub
Public Sub Add_ToPackage(ByVal lngPackageID As Long)

    If Check_AddedToPackage = False Then

        DoCmd.SetWarnings False
        
        DoCmd.RunSQL "INSERT INTO tbl_Package_Clas"
        
        DoCmd.SetWarnings True

    End If

End Sub
Public Function Check_AddedToPackage(ByVal lngPackageID As Long) As Boolean

    if isnull(DLookup("ID"

End Function