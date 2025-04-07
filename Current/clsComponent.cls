Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum enuComponent_Types

    ctTable
    ctQuery
    ctForm
    ctModule
    ctClassModule
    
End Enum

Private lngID As Long
Private strName As String
Private strDescription As String
Private lngType_FK As Long
Private lngApplication_Part As Long
Private datCreateTS As Date
Private datLastScanTS As Date
Private datLastUpdateTS As Date
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
Public Property Get Type_FK() As Long

    Type_FK = lngType_FK

End Property
Public Property Let Type_FK(lngType_FK_1 As Long)

    lngType_FK = lngType_FK_1

End Property
Public Property Get Application_Part() As Long

    Application_Part = lngApplication_Part

End Property
Public Property Let Application_Part(lngApplication_Part_1 As Long)

    lngApplication_Part = lngApplication_Part_1

End Property
Public Property Get CreateTS() As Date

    CreateTS = datCreateTS

End Property
Public Property Let CreateTS(datCreateTS_1 As Date)

    datCreateTS = datCreateTS_1

End Property
Public Property Get LastScanTS() As Date

    LastScanTS = datLastScanTS

End Property
Public Property Let LastScanTS(datLastScanTS_1 As Date)

    datLastScanTS = datLastScanTS_1

End Property
Public Property Get LastUpdateTS() As Date

    LastUpdateTS = datLastUpdateTS

End Property
Public Property Let LastUpdateTS(datLastUpdateTS_1 As Date)

    datLastUpdateTS = datLastUpdateTS_1

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
    Me.Type_FK = 0
    Me.Application_Part = 0
    Me.CreateTS = 0
    Me.LastScanTS = 0
    Me.LastUpdateTS = 0
    Me.Active = False

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If Isnull(dlookup("ID", "tbl_Component", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Component WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Component (Name, Description, Type_FK, Application_Part_FK, CreateTS, LastScanTS, LastUpdateTS, Active) VALUES (, , , , , , , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Component")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Component WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strName = rs![Name]
        strDescription = rs![Description]
        lngType_FK = rs![Type_FK]
        lngApplication_Part = rs![Application_Part]
        datCreateTS = rs![CreateTS]
        datLastScanTS = rs![LastScanTS]
        datLastUpdateTS = rs![LastUpdateTS]
        blnActive = rs![Active]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component SET Name = , Description = , Type_FK = , Application_Part_FK = , CreateTS = , LastScanTS = , LastUpdateTS = , Active =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub