Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private apApplication_Part As New clsApplication_Part


'Component
Private strName As String
Private strDescription As String
Private lngType_FK As Long
Private lngApplication_Part As Long
Private datCreateTS As Date
Private datLastScanTS As Date
Private datLastUpdateTS As Date
Private blnActive As Boolean

'Module
Private lngID As Long
Private strCode As String
Private lngLOC As Long
Private blnIsClassModule As Boolean
Private strInstancing As String
Private lngComponent_FK As Long

'Component
Public Property Get Name() As String

    Name = Application_Part.Name

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
Public Property Get Application_Part() As clsApplication_Part

    Application_Part = apApplication_Part

End Property
Public Property Set Application_Part(lngApplication_Part_1 As clsApplication_Part)

    Set apApplication_Part = lngApplication_Part_1

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

'Module
Public Property Get ID() As Long

    ID = lngID

End Property
Public Property Let ID(lngID_1 As Long)

    lngID = lngID_1

End Property
Public Property Get Code() As String

    Code = strCode

End Property
Public Property Let Code(strCode_1 As String)

    strCode = strCode_1

End Property
Public Property Get LOC() As Long

    LOC = lngLOC

End Property
Public Property Let LOC(lngLOC_1 As Long)

    lngLOC = lngLOC_1

End Property
Public Property Get IsClassModule() As Boolean

    IsClassModule = blnIsClassModule

End Property
Public Property Let IsClassModule(blnIsClassModule_1 As Boolean)

    blnIsClassModule = blnIsClassModule_1

End Property
Public Property Get Instancing() As String

    Instancing = strInstancing

End Property
Public Property Let Instancing(strInstancing_1 As String)

    strInstancing = strInstancing_1

End Property
Public Property Get Component_FK() As Long

    Component_FK = lngComponent_FK

End Property
Public Property Let Component_FK(lngComponent_FK_1 As Long)

    lngComponent_FK = lngComponent_FK_1

End Property
Public Sub Clear()

    Me.ID = 0
    Me.Code = ""
    Me.LOC = 0
    Me.IsClassModule = False
    Me.Instancing = ""
    Me.Component_FK = 0

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Module SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If IsNull(dlookup("ID", "tbl_Component_Module", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Module SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Component_Module WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Component_Module (Code, LOC, IsClassModule, Instancing, Component_FK) VALUES (, , , , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Component_Module")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Component_Module WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strCode = rs![Code]
        lngLOC = rs![LOC]
        blnIsClassModule = rs![IsClassModule]
        strInstancing = rs![Instancing]
        lngComponent_FK = rs![Component_FK]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Module SET Code = , LOC = , IsClassModule = , Instancing = , Component_FK =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub