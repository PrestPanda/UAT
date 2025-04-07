Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private lngID As Long
Private strName As String
Private blnIsAddedInCode As Boolean
Private strCode As String
Private lngClass_FK As Long
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
Public Property Get IsAddedInCode() As Boolean

    IsAddedInCode = blnIsAddedInCode

End Property
Public Property Let IsAddedInCode(blnIsAddedInCode_1 As Boolean)

    blnIsAddedInCode = blnIsAddedInCode_1

End Property
Public Property Get Code() As String

    Code = strCode

End Property
Public Property Let Code(strCode_1 As String)

    strCode = strCode_1

End Property
Public Property Get Class_FK() As Long

    Class_FK = lngClass_FK

End Property
Public Property Let Class_FK(lngClass_FK_1 As Long)

    lngClass_FK = lngClass_FK_1

End Property
Public Property Get Active() As Boolean

    Active = blnActive

End Property
Public Property Let Active(blnActive_1 As Boolean)

    blnActive = blnActive_1

End Property
Public Sub Clear_()

    Me.ID = 0
    Me.Name = ""
    Me.IsAddedInCode = False
    Me.Code = ""
    Me.Class_FK = 0
    Me.Active = False

End Sub


Public Sub DB_Activate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Class_Property SET Active = true WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Function DB_Check() As Boolean

    If Isnull(dlookup("ID", "tbl_Class_Property", "ID =" & lngID)) = False Then
        DB_Check = True
    Else
        DB_Check = False
    End If

End Function
Public Sub DB_Deactivate()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Class_Property SET Active = false WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Delete()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "DELETE * FROM tbl_Class_Property WHERE ID =" & lngID
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Insert()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "INSERT INTO tbl_Class_Property (Name, IsAddedInCode, Code, Class_FK, Active) VALUES ('" & strName & "', 5, '" & strCode & "', lngClass_FK, 5)"
    
    DoCmd.setwarnings True

End Sub
Public Sub DB_Load()

    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT * FROM tbl_Class_Property WHERE ID = " & lngID
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)
    If Not rs.EOF Then
        lngID = rs![ID]
        strName = rs![Name]
        blnIsAddedInCode = rs![IsAddedInCode]
        strCode = rs![Code]
        lngClass_FK = rs![Class_FK]
        blnActive = rs![Active]
    End If
    rs.Close
    Set rs = Nothing
    

End Sub
Public Sub DB_Update()

    DoCmd.setwarnings False
    
    DoCmd.RunSQL "UPDATE tbl_Class_Property SET Name = '" & strName & "', IsAddedInCode = 5, Code = '" & strCode & "', Class_FK = lngClass_FK, Active = 5 WHERE ID = " & lngID
    
    DoCmd.setwarnings True

End Sub