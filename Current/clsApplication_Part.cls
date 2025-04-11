Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private lngID As Long
Private strName As String
Private strDescription As String
Private strPath As String
Private lngType_Programm_FK As Long
Private lngType_Functional_FK As Long
Private lngApplication_FK As Long
Private lngLOC As Long
Private lngTable_Count As Long
Private lngQuery_Count As Long
Private lngForm_Count As Long
Private lngForm_LOC As Long
Private lngModule_Count As Long
Private lngModule_LOC As Long
Private lngClassModule_Count As Long
Private lngClassModule_LOC As Long
Private blnActive As Boolean
Private datLastScanTS As Date

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
Public Property Get Description() As String

    Description = strDescription

End Property
Public Property Let Description(strDescription_1 As String)

    strDescription = strDescription_1

End Property
Public Property Get Path() As String

    Path = strPath

End Property
Public Property Let Path(strPath_1 As String)

    strPath = strPath_1

End Property
Public Property Get Type_Programm_FK() As Long

    Type_Programm_FK = lngType_Programm_FK

End Property
Public Property Let Type_Programm_FK(lngType_Programm_FK_1 As Long)

    lngType_Programm_FK = lngType_Programm_FK_1

End Property
Public Property Get Type_Functional_FK() As Long

    Type_Functional_FK = lngType_Functional_FK

End Property
Public Property Let Type_Functional_FK(lngType_Functional_FK_1 As Long)

    lngType_Functional_FK = lngType_Functional_FK_1

End Property
Public Property Get Application_FK() As Long

    Application_FK = lngApplication_FK

End Property
Public Property Let Application_FK(lngApplication_FK_1 As Long)

    lngApplication_FK = lngApplication_FK_1

End Property
Public Property Get LOC() As Long

    LOC = lngLOC

End Property
Public Property Let LOC(lngLOC_1 As Long)

    lngLOC = lngLOC_1

End Property
Public Property Get Table_Count() As Long

    Table_Count = lngTable_Count

End Property
Public Property Let Table_Count(lngTable_Count_1 As Long)

    lngTable_Count = lngTable_Count_1

End Property
Public Property Get Query_Count() As Long

    Query_Count = lngQuery_Count

End Property
Public Property Let Query_Count(lngQuery_Count_1 As Long)

    lngQuery_Count = lngQuery_Count_1

End Property
Public Property Get Form_Count() As Long

    Form_Count = lngForm_Count

End Property
Public Property Let Form_Count(lngForm_Count_1 As Long)

    lngForm_Count = lngForm_Count_1

End Property
Public Property Get Form_LOC() As Long

    Form_LOC = lngForm_LOC

End Property
Public Property Let Form_LOC(lngForm_LOC_1 As Long)

    lngForm_LOC = lngForm_LOC_1

End Property
Public Property Get Module_Count() As Long

    Module_Count = lngModule_Count

End Property
Public Property Let Module_Count(lngModule_Count_1 As Long)

    lngModule_Count = lngModule_Count_1

End Property
Public Property Get Module_LOC() As Long

    Module_LOC = lngModule_LOC

End Property
Public Property Let Module_LOC(lngModule_LOC_1 As Long)

    lngModule_LOC = lngModule_LOC_1

End Property
Public Property Get ClassModule_Count() As Long

    ClassModule_Count = lngClassModule_Count

End Property
Public Property Let ClassModule_Count(lngClassModule_Count_1 As Long)

    lngClassModule_Count = lngClassModule_Count_1

End Property
Public Property Get ClassModule_LOC() As Long

    ClassModule_LOC = lngClassModule_LOC

End Property
Public Property Let ClassModule_LOC(lngClassModule_LOC_1 As Long)

    lngClassModule_LOC = lngClassModule_LOC_1

End Property
Public Property Get Active() As Boolean

    Active = blnActive

End Property
Public Property Let Active(blnActive_1 As Boolean)

    blnActive = blnActive_1

End Property
Public Property Get LastScanTS() As Date

    LastScanTS = datLastScanTS

End Property
Public Property Let LastScanTS(datLastScanTS_1 As Date)

    datLastScanTS = datLastScanTS_1

End Property
Public Sub Clear()

    Me.ID = 0
    Me.Name = ""
    Me.Description = ""
    Me.Path = ""
    Me.Type_Programm_FK = 0
    Me.Type_Functional_FK = 0
    Me.Application_FK = 0
    Me.LOC = 0
    Me.Table_Count = 0
    Me.Query_Count = 0
    Me.Form_Count = 0
    Me.Form_LOC = 0
    Me.Module_Count = 0
    Me.Module_LOC = 0
    Me.ClassModule_Count = 0
    Me.ClassModule_LOC = 0
    Me.Active = False
    Me.LastScanTS = 0

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application_Part SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If IsNull(dlookup("ID", "tbl_Application_Part", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application_Part SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Application_Part WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Application_Part (Name, Description, Path, Type_Programm_FK, Type_Functional_FK, Application_FK, LOC, Table_Count, Query_Count, Form_Count, Form_LOC, Module_Count, Module_LOC, ClassModule_Count, ClassModule_LOC, Active, LastScanTS) VALUES (, , , , , , , , , , , , , , , , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Application_Part")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Application_Part WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strName = rs![Name]
        strDescription = rs![Description]
        strPath = rs![Path]
        lngType_Programm_FK = rs![Type_Programm_FK]
        lngType_Functional_FK = rs![Type_Functional_FK]
        lngApplication_FK = rs![Application_FK]
        lngLOC = rs![LOC]
        lngTable_Count = rs![Table_Count]
        lngQuery_Count = rs![Query_Count]
        lngForm_Count = rs![Form_Count]
        lngForm_LOC = rs![Form_LOC]
        lngModule_Count = rs![Module_Count]
        lngModule_LOC = rs![Module_LOC]
        lngClassModule_Count = rs![ClassModule_Count]
        lngClassModule_LOC = rs![ClassModule_LOC]
        blnActive = rs![Active]
        datLastScanTS = rs![LastScanTS]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application_Part SET Name = , Description = , Path = , Type_Programm_FK = , Type_Functional_FK = , Application_FK = , LOC = , Table_Count = , Query_Count = , Form_Count = , Form_LOC = , Module_Count = , Module_LOC = , ClassModule_Count = , ClassModule_LOC = , Active = , LastScanTS =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub