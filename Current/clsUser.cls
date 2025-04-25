Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

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

    If DB_Check = True Then
        Exists = True
    Else
        Exists = False
    End If

End Function
Public Sub Delete()

    If Exists = True Then
        DB_Delete
    End If

End Sub
Public Sub LoadByID()

    lngID = CLng(InputBox("Bitte ID eingeben", "Datensatz laden"))
    If lngID > 0 Then
        Load
    End If

End Sub
Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_User SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If IsNull(DLookup("ID", "tbl_User", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_User SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_User WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_User " & _
    "(Name, CreateTS, CreateUser_FK, LastUpdateTS, LastUpdateUser_FK, IsActive, LastActiveTS, LastActiveUser_FK) " & _
    "VALUES ('" & strName & "', #CreateTS#, CreateUser_FK, #LastUpdateTS#, LastUpdateUser_FK, 5, #LastActiveTS#, LastActiveUser_FK)"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_User")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_User WHERE ID = " & lngID
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
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_User SET Name = strName, CreateTS = datCreateTS, CreateUser_FK = lngCreateUser_FK, LastUpdateTS = datLastUpdateTS, LastUpdateUser_FK = lngLastUpdateUser_FK, IsActive = blnIsActive, LastActiveTS = datLastActiveTS, LastActiveUser_FK = lngLastActiveUser_FK WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub