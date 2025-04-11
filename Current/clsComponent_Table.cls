Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private apApplication_Part As New clsApplication_Part

'Component


'Table
Private lngID As Long
Private lngColumnCount As Long
Private lngSizeKB As Long
Private blnIsSystemTable As Boolean
Private blnIsLinkedTable As Boolean
Private strLinkedTable_Path As String
Private lngComponent_FK As Long

'Component
'To-Do:
Public Property Get Name() As String

    Name = apApplication_Part.Name

End Property
Public Property Let Name(strName_1 As String)

    apApplication_Part.Name = strName_1

End Property
Public Property Get Description() As String

    Description = apApplication_Part.Description

End Property
Public Property Let Description(strDescription_1 As String)

    apApplication_Part.Description = strDescription_1

End Property
Public Property Get Type_FK() As Long

    Type_FK = lngType_FK

End Property
Public Property Let Type_FK(lngType_FK_1 As Long)

    lngType_FK = lngType_FK_1

End Property
Public Property Get Application_Part() As clsApplication_Part

    Application_Part = apApplication_Part

End Property
Public Property Set Application_Part(lngApplication_Part_1 As clsApplication_Part)

    Set lngApplication_Part = lngApplication_Part_1

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

'Table
Public Property Get ID() As Long

    ID = lngID

End Property
Public Property Let ID(lngID_1 As Long)

    lngID = lngID_1

End Property
Public Property Get ColumnCount() As Long

    ColumnCount = lngColumnCount

End Property
Public Property Let ColumnCount(lngColumnCount_1 As Long)

    lngColumnCount = lngColumnCount_1

End Property
Public Property Get SizeKB() As Long

    SizeKB = lngSizeKB

End Property
Public Property Let SizeKB(lngSizeKB_1 As Long)

    lngSizeKB = lngSizeKB_1

End Property
Public Property Get IsSystemTable() As Boolean

    IsSystemTable = blnIsSystemTable

End Property
Public Property Let IsSystemTable(blnIsSystemTable_1 As Boolean)

    blnIsSystemTable = blnIsSystemTable_1

End Property
Public Property Get IsLinkedTable() As Boolean

    IsLinkedTable = blnIsLinkedTable

End Property
Public Property Let IsLinkedTable(blnIsLinkedTable_1 As Boolean)

    blnIsLinkedTable = blnIsLinkedTable_1

End Property
Public Property Get LinkedTable_Path() As String

    LinkedTable_Path = strLinkedTable_Path

End Property
Public Property Let LinkedTable_Path(strLinkedTable_Path_1 As String)

    strLinkedTable_Path = strLinkedTable_Path_1

End Property
Public Property Get Component_FK() As Long

    Component_FK = lngComponent_FK

End Property
Public Property Let Component_FK(lngComponent_FK_1 As Long)

    lngComponent_FK = lngComponent_FK_1

End Property
Public Sub Clear()

    Me.ID = 0
    Me.ColumnCount = 0
    Me.SizeKB = 0
    Me.IsSystemTable = False
    Me.IsLinkedTable = False
    Me.LinkedTable_Path = ""
    Me.Component_FK = 0

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Table SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If IsNull(dlookup("ID", "tbl_Component_Table", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Table SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Component_Table WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Component_Table (ColumnCount, SizeKB, IsSystemTable, IsLinkedTable, LinkedTable_Path, Component_FK) VALUES (, , , , , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Component_Table")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Component_Table WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        lngColumnCount = rs![ColumnCount]
        lngSizeKB = rs![SizeKB]
        blnIsSystemTable = rs![IsSystemTable]
        blnIsLinkedTable = rs![IsLinkedTable]
        strLinkedTable_Path = rs![LinkedTable_Path]
        lngComponent_FK = rs![Component_FK]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Table SET ColumnCount = , SizeKB = , IsSystemTable = , IsLinkedTable = , LinkedTable_Path = , Component_FK =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub