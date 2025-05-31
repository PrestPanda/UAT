Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim clsLog As New clsLog
Dim clsCoding_SQL As New clsCoding_SQL



Private lngID As Long
Private strName As String
Private datCreateTS As Date
Private lngCreateUser_FK As Long
Private datLastUpdateTS As Date
Private lngLastUpdateUser_FK As Long
Private blnIsActive As Boolean
Private datLastActiveTS As Date
Private lngLastActiveUser_FK As Long

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
    
        If IsNull(DLookup("ID", "tbl_Table", "ID =" & lngID)) = False Then
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
    
    DoCmd.RunSQL "DELETE * FROM tbl_Table WHERE ID =" & lngID
    
    DoCmd.SetWarnings True
    
    Reset

End Sub
Private Sub DB_Insert()

    If IsNull(DMax("ID", "tbl_Table")) = False Then
        lngID = DMax("ID", "tbl_Table") + 1
    End If
    
    datCreateTS = Now()
    datLastUpdateTS = Now()
    blnIsActive = True
    
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Table " & _
        "(ID, Name, CreateTS, CreateUser_FK, LastUpdateTS, LastUpdateUser_FK, IsActive, LastActiveTS, LastActiveUser_FK) " & _
        "VALUES (" & lngID & ", '" & strName & "', " & Coding_SQL.Get_CDate_ByDate_AsString(datCreateTS) & ", " & lngCreateUser_FK & ", " & Coding_SQL.Get_CDate_ByDate_AsString(datLastUpdateTS) & ", " & lngLastUpdateUser_FK & ", " & CInt(blnIsActive) & ", " & Coding_SQL.Get_CDate_ByDate_AsString(datLastActiveTS) & ", " & lngLastActiveUser_FK & ")"
    
    DoCmd.SetWarnings True
    
    lngID = DMax("ID", "tbl_Table")
    

End Sub
Private Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    If DB_Check = True Then
    
        strSQL = "SELECT * FROM tbl_Table WHERE ID = " & lngID
    
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
        End If
    
        rs.Close
        Set rs = Nothing
    
    End If
    

End Sub
Private Sub DB_Update()

    datLastUpdateTS = Now()
    
    DoCmd.SetWarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Table SET Name = '" & strName & "', CreateTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datCreateTS) & ", CreateUser_FK = " & lngCreateUser_FK & ", LastUpdateTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datLastUpdateTS) & ", LastUpdateUser_FK = " & lngLastUpdateUser_FK & ", IsActive = " & CInt(blnIsActive) & ", LastActiveTS = " & Coding_SQL.Get_CDate_ByDate_AsString(datLastActiveTS) & ", LastActiveUser_FK = " & lngLastActiveUser_FK & " WHERE ID = " & lngID
    
    DoCmd.SetWarnings True

End Sub
Private Sub GetObject_FromName()



End Sub