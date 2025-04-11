Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private lngID As Long
Private strName As String
Private strDescription As String
Private blnActive As Boolean

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
Public Property Get Active() As Boolean

    Active = blnActive

End Property
Public Property Let Active(blnActive_1 As Boolean)

    blnActive = blnActive_1

End Property
Public Sub Clear()

    Me.ID = 0
    Me.Name = ""
    Me.Description = ""
    Me.Active = False

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If IsNull(dlookup("ID", "tbl_Application", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Application WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Application (Name, Description, Active) VALUES (, , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Application")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Application WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strName = rs![Name]
        strDescription = rs![Description]
        blnActive = rs![Active]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Application SET Name = , Description = , Active =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub