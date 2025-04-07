Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Component
Private strName As String
Private strDescription As String
Private lngType_FK As Long
Private lngApplication_Part As Long
Private datCreateTS As Date
Private datLastScanTS As Date
Private datLastUpdateTS As Date
Private blnActive As Boolean

'Form
Private lngID As Long
Private strCode As String
Private lngLOC As Long
Private lngControl_Count As Long
Private blnSetting_AutoCenter As Boolean
Private blnSetting_NavigationPanes As Boolean
Private blnSetting_PopUp As Boolean
Private blnSetting_ScrollBars As Boolean
Private strSetting_StandardView As String
Private lngComponent_FK As Long
'Component
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

'Form
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
Public Property Get Control_Count() As Long

    Control_Count = lngControl_Count

End Property
Public Property Let Control_Count(lngControl_Count_1 As Long)

    lngControl_Count = lngControl_Count_1

End Property
Public Property Get Setting_AutoCenter() As Boolean

    Setting_AutoCenter = blnSetting_AutoCenter

End Property
Public Property Let Setting_AutoCenter(blnSetting_AutoCenter_1 As Boolean)

    blnSetting_AutoCenter = blnSetting_AutoCenter_1

End Property
Public Property Get Setting_NavigationPanes() As Boolean

    Setting_NavigationPanes = blnSetting_NavigationPanes

End Property
Public Property Let Setting_NavigationPanes(blnSetting_NavigationPanes_1 As Boolean)

    blnSetting_NavigationPanes = blnSetting_NavigationPanes_1

End Property
Public Property Get Setting_PopUp() As Boolean

    Setting_PopUp = blnSetting_PopUp

End Property
Public Property Let Setting_PopUp(blnSetting_PopUp_1 As Boolean)

    blnSetting_PopUp = blnSetting_PopUp_1

End Property
Public Property Get Setting_ScrollBars() As Boolean

    Setting_ScrollBars = blnSetting_ScrollBars

End Property
Public Property Let Setting_ScrollBars(blnSetting_ScrollBars_1 As Boolean)

    blnSetting_ScrollBars = blnSetting_ScrollBars_1

End Property
Public Property Get Setting_StandardView() As String

    Setting_StandardView = strSetting_StandardView

End Property
Public Property Let Setting_StandardView(strSetting_StandardView_1 As String)

    strSetting_StandardView = strSetting_StandardView_1

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
    Me.Control_Count = 0
    Me.Setting_AutoCenter = False
    Me.Setting_NavigationPanes = False
    Me.Setting_PopUp = False
    Me.Setting_ScrollBars = False
    Me.Setting_StandardView = ""
    Me.Component_FK = 0

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Form SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If Isnull(dlookup("ID", "tbl_Component_Form", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Form SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Component_Form WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Component_Form (Code, LOC, Control_Count, Setting_AutoCenter, Setting_NavigationPanes, Setting_PopUp, Setting_ScrollBars, Setting_StandardView, Component_FK) VALUES (, , , , , , , , )"
    
    DoCmd.setwarnings True
    
    lngID = DMax("ID", "tbl_Component_Form")
    

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Component_Form WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strCode = rs![Code]
        lngLOC = rs![LOC]
        lngControl_Count = rs![Control_Count]
        blnSetting_AutoCenter = rs![Setting_AutoCenter]
        blnSetting_NavigationPanes = rs![Setting_NavigationPanes]
        blnSetting_PopUp = rs![Setting_PopUp]
        blnSetting_ScrollBars = rs![Setting_ScrollBars]
        strSetting_StandardView = rs![Setting_StandardView]
        lngComponent_FK = rs![Component_FK]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Component_Form SET Code = , LOC = , Control_Count = , Setting_AutoCenter = , Setting_NavigationPanes = , Setting_PopUp = , Setting_ScrollBars = , Setting_StandardView = , Component_FK =  WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub