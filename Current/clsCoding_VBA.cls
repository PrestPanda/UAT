Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
'Die Coding_VBA stellt Funktionalitäten zur automatisierten Erstellung und Bearbeitung von VBA COde bereit


Private Coding_SQL As New clsCoding_SQL
Private Coding As New clsCoding
Private Coding_VBA As New clsCoding_VBA
Private Format As New clsFormat
Private Log As New clsLog


Public Function Get_Code_ClassModule(strClassName As String, _
    Properties() As Variant, _
    Methods() As Variant) As String
'Refactored

    Dim strCode As String
    Dim strTableName As String
    Dim strClassModuleName As String
    Dim lngCounter As Long
    
    strTableName = "tbl_" & strClassName
    strClassModuleName = "cls" & strClassName
  
    
    strCode = Get_Code_ClassModule_PrivateVariables(Properties()) & vbNewLine
    strCode = strCode & Get_Code_ClassModule_Properties(Properties()) & vbNewLine
    strCode = strCode & Get_Code_Module_Methods(strTableName, Properties(), Methods()) & vbNewLine

    Get_Code_Module = strCode

End Function
'-------------------------------------------- CODE - Module - PRIVATE VARIABLES -----------------------------------
Public Function Get_Code_ClassModule_PrivateVariables(Properties() As Variant) As String
'Refactored

    Log.WriteLine "Code für private Variablen wird generiert."
    
    Dim strCode As String

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If strCode = "" Then
        
            strCode = strCode & _
                Coding_VBA.Get_Code_PrivateVariable_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2))
        Else
        
            strCode = strCode & vbNewLine & _
                Coding_VBA.Get_Code_PrivateVariable_AsString_ByString _
                    (Properties(lngCounter, 1), _
                    Properties(lngCounter, 2))
        End If
   
    Next
    
    Get_Code_ClassModule_PrivateVariables = strCode
    
    Log.WriteLine "Code generiert."
    

End Function
Public Function Get_Code_PrivateVariable_AsString_ByString(ByVal strName As String, _
ByVal strDataType As String) As String
'Refactored

    Get_Code_PrivateVariable_AsString_ByString = _
        Coding.Get_Variable_DimensionStatement_ByString( _
            vPrivate, strName, strDataType)

End Function
'-------------------------------------------- CODE - Module - PROPERTIES -----------------------------------
Public Function Get_Code_ClassModule_Properties(Properties() As Variant) As String
'Refactored
    Dim strCode As String
    
    Log.WriteLine "Code für Properties wird generiert"
    
   For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If strCode = "" Then
            strCode = strCode & _
                Get_Code_Module_Property_ByString( _
                Properties(lngCounter, 1), _
                Properties(lngCounter, 2), "Public")
        Else
            strCode = strCode & vbNewLine & _
                Get_Code_Module_Property_ByString( _
                Properties(lngCounter, 1), _
                Properties(lngCounter, 2), "Public")
        End If
      
    
    Next
    
    Log.WriteLine "Code generiert."
    
    Get_Code_ClassModule_Properties = strCode
    

End Function
Public Function Get_Code_Module_Property_ByString(ByVal strPropertyName As String, _
    ByVal strDataType As String, _
    ByVal strVisibility As String) As String
'Refactored

    Dim strCode As String

    strCode = Get_Code_Module_Property_Get_ByString(strPropertyName, strDataType, strVisibility)
    strCode = strCode & vbNewLine & Get_Code_Module_Property_Let_ByString(strPropertyName, strDataType, strVisibility)
                            
    Get_Code_Module_Property = strCode
            
End Function
Public Function Get_Code_Module_Property_Get_ByString(strPropertyName As String, _
    strDataType As String, _
    strVisibility As String) As String
'Refactored

    Get_Code_Module_Property_Get_ByString = strVisibility & _
                            " Property Get " & strPropertyName & "() As " & _
                             vbNewLine & vbNewLine & _
                             vbTab & strPropertyName & " = " & _
                             Coding.Get_Variable_Name_AsString_ByString(strPropertyName, strDataType) & _
                             vbNewLine & vbNewLine & _
                            "End Property"
    
End Function
Public Function Get_Code_Module_Property_Let_ByString(strPropertyName As String, _
    strDataType As String, _
    strVisibility As String) As String
'Refactored

    Get_Code_Module_Property_Let_ByString = strVisibility & _
                            " Property Let " & strPropertyName & "(" & _
                            Coding.Get_Variable_Name_AsString_ByString(strPropertyName, strDataType) & "_1 as " & _
                            strDataType & ")" & _
                             vbNewLine & vbNewLine & _
                             vbTab & Coding.Get_Variable_Name_AsString_ByString(strPropertyName, strDataType) & " = " & _
                             Coding.Get_Variable_Name_AsString_ByString(strPropertyName, strDataType) & "_1" & _
                             vbNewLine & vbNewLine & _
                            "End Property"
    
End Function
'-------------------------------------------- CODE - Module - METHODS -----------------------------------
Public Function Get_Code_Module_Methods(strTableName As String, _
    Properties() As Variant, Methods() As Variant) As String
'In Refactoring
    Dim strCode As String
    Dim strCodeAddition As String
    Dim lngCounter As Long
    
    Log.WriteLine "Code für Methods wird generiert"
    
   For lngCounter = LBound(Methods()) To UBound(Methods())
    
        If IsNull(DLookup("FunctionName", "110_tblClassBuilder_Method_Draft", _
            "Name = '" & Methods(lngCounter, 1) & "'")) = False Then
            
            'Code der Standardmethode in den Wrap einfügen
'            strCodeAddition = Get_Code_Module_Method_Draft( _
'                strTableName, _
'                DLookup("FunctionName", "110_tblClassBuilder_Method_Draft", _
'                "Name = '" & Methods(lngCounter, 1) & "'"))
            
        Else
        
            'Nur den Wrap erstellen
            'To-Do: Hier weiterarbeiten
'            strCodeAddition = Get_Code_Wrap_VBA_Method
            
        End If
    
        If strCode = "" Then
            strCode = strCodeAddition
        Else
            strCode = strCode & vbNewLine & strCodeAddition
        End If
    
    Next
    
    Log.WriteLine "Code generiert."
    
    Get_Code_Module_Methods = strCode

End Function
'-------------------------------------------- DRAFT SELECT -----------------------------------
Public Function Get_Code_Module_Method_Draft(strTableName As String, _
    strDraftName As String, Properties() As Variant) As String
'Refactored

    Select Case DLookup("Package_PK", "110_tblCLassBuilder_Method_Draft", "Name = '" & strDraftName & "'")
    
        Case DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & "DB" & "'")
            Get_Code_Module_Method_Draft = _
                Get_Code_Module_Method_Draft_DB(strTableName, strDraftName, Properties())
        
        Case DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & "Standard" & "'")
            Get_Code_Module_Method_Draft = _
                Get_Code_Module_Method_Draft_Standard(strDraftName, Properties())
                
        Case DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & "Tracking" & "'")
        
        
        Case DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & "DEV" & "'")
        
    
    End Select

End Function
'-------------------------------------------- DRAFT Standard -----------------------------------
Public Function Get_Code_Module_Method_Draft_Standard(strDraftName As String, Properties() As Variant) As String
'Refactored

    Select Case strDraftName
    
        Case "Reset"
            Get_Code_Module_Method_Draft_Standard = _
                Get_Code_Module_Method_Draft_Standard_Reset(Properties())
        
    
    End Select

End Function
Public Function Get_Code_Module_Method_Draft_Standard_Reset(Properties() As Variant) As String
'Refactored

    Dim strCode As String
    Dim lngCounter As Long
    Dim strDefaultValue As String

    strCode = "Public Sub Reset" & "()" & vbCrLf & vbCrLf


    For lngCounter = LBound(Properties()) To UBound(Properties())
    
    
        Select Case Coding.Get_Variable_Type_AsEnum_ByString(Properties(lngCounter, 2))
        
            Case vtString, vtStringLong
                strDefaultValue = """"""
                
            Case vtLong, vtDouble, vtAuto
                strDefaultValue = "0"
                
            Case vtBoolean
                strDefaultValue = "False"
                
            Case vtDate
                strDefaultValue = "0"
                
            Case Else
                strDefaultValue = "Nothing"
                
        End Select

        strCode = strCode & "    Me." & Coding.Get_Variable_Name_AsString_ByString( _
            Properties(lngCounter, 1), _
            Properties(lngCounter, 2)) & " = " & strDefaultValue & vbCrLf
            
        
    Next lngCounter

    strCode = strCode & vbCrLf & "End Sub" & vbCrLf

    Get_Code_Module_Method_Draft_Standard_Reset = strCode

End Function
'-------------------------------------------- DRAFT DB -----------------------------------
Public Function Get_Code_Module_Method_Draft_DB(strTableName As String, _
    strDraftName As String, Properties() As Variant) As String
'Refactored

    Select Case strDraftName
    
        Case "DB_Activate"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Activate(strTableName)
        
        Case "DB_Check"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Check(strTableName)
        
        Case "DB_Deactivate"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Deactivate(strTableName)
        
        Case "DB_Delete"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Delete(strTableName)
        
        Case "DB_Insert"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Insert(strTableName, Properties())
        
        Case "DB_Load"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Load(strTableName, Properties())
        
        Case "DB_Update"
            Get_Code_Module_Method_Draft_DB = Get_Code_Module_Method_Draft_DB_Update(strTableName, Properties())
    
    End Select

End Function
Public Function Get_Code_Module_Method_Draft_DB_Activate(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Activate = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Activate", "", _
                                            Get_Code_Wrap_VBA_SQL_Statement( _
                                            Coding_SQL.Get_DB_Activate(strTableName)) _
                                            , vtNoType)

End Function
Public Function Get_Code_Module_Method_Draft_DB_Check(strTableName As String) As String
'Refactored

    Dim strCode As String

    strCode = "If Isnull(dlookup(" & """ID""" & ",""" & strTableName & """," & """ID =""" & " & lngID)) = false then" & vbNewLine & _
        vbTab & "DB_Check = true" & vbNewLine & _
        "Else" & vbNewLine & _
        vbTab & "DB_Check = false" & vbNewLine & _
        "End If"

    Get_Code_Module_Method_Draft_DB_Check = _
        Get_Code_Wrap_VBA_Method(vPublic, mtFunction, "DB_Check", "", strCode, vtBoolean)

End Function
Public Function Get_Code_Module_Method_Draft_DB_Deactivate(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Deactivate = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Deactivate", "", _
            Get_Code_Wrap_VBA_SQL_Statement( _
            Coding_SQL.Get_DB_Deactivate(strTableName)) _
            , vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Delete(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Delete = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Delete", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Delete(strTableName)) _
                                        , vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Insert(strTableName As String, Properties() As Variant) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Insert = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Insert", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Insert(strTableName, _
                                        Properties())) & _
                                        Get_Code_Module_Method_DB_Insert_UpdateID(strTableName) _
                                        , vtNoType)

End Function
Public Function Get_Code_Module_Method_DB_Insert_UpdateID(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_DB_Insert_UpdateID = vbNewLine & vbNewLine & _
                                        "lngID = DMax(""" & "ID""" & ", """ & strTableName & """)" & _
                                        vbNewLine

End Function
Public Function Get_Code_Module_Method_Draft_DB_Load(strTableName As String, Properties() As Variant) As String
'Refactored

    Dim strCode As String
    Dim lngCounter As Long
    Dim strVarName As String

    strCode = "Dim rs As DAO.Recordset" & vbCrLf
    strCode = strCode & "Dim strSQL As String" & vbCrLf
    strCode = strCode & "strSQL = ""SELECT * FROM " & strTableName & " WHERE ID = "" & lngID" & vbCrLf
    strCode = strCode & "Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)" & vbCrLf
    strCode = strCode & "If Not rs.EOF Then" & vbCrLf

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        strVarName = Coding.Get_Variable_Name_AsString_ByString( _
            Properties(lngCounter, 1), _
            Properties(lngCounter, 2))
            
        strCode = strCode & "    " & strVarName & " = rs![" & Properties(lngCounter, 1) & "]" & vbCrLf
        
    Next lngCounter

    strCode = strCode & "End If" & vbCrLf
    strCode = strCode & "rs.Close" & vbCrLf
    strCode = strCode & "Set rs = Nothing" & vbCrLf


     Get_Code_Module_Method_Draft_DB_Load = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Load", "", strCode, vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Update(strTableName As String, Properties() As Variant) As String
'Refactored

    Get_Code_Module_Method_DB_Update = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Update", "", _
        Get_Code_Wrap_VBA_SQL_Statement( _
        Coding_SQL.Get_DB_Update(strTableName, Properties)) _
            , vtNoType)


End Function
'--------------------------------------------Code Wrap -----------------------------------
Public Function Get_Code_Wrap_VBA_SQL_Statement(strSQLStatement As String) As String

    Get_Code_Wrap_VBA_SQL_Statement = "docmd.setwarnings false" & vbNewLine & vbNewLine & _
                        "docmd.runsql """ & strSQLStatement & vbNewLine & vbNewLine & _
                        "docmd.setwarnings true"

End Function
Public Function Get_Code_Wrap_VBA_Method(Visability As enuCoding_Visibility, vtType As enuCoding_Method_Types, _
    strMethodName As String, strParameters As String, strVBACode As String, vtReturnType As enuCoding_DataTypes)
    
    Dim strCode As String
   

    If vtReturnType = vtNoType Then
        strCode = Coding.Get_Visibility_AsString_ByEnum(Visability) & " " & Coding.Get_Method_Type_AsString_ByEnum(vtType) & " " & _
            strMethodName & "(" & strParameters & ")"
    Else
        strCode = Coding.Get_Visibility_AsString_ByEnum(Visability) & " " & Coding.Get_Method_Type_AsString_ByEnum(vtType) & " " & _
            strMethodName & "(" & strParameters & ") as " & Coding.Get_Variable_Type_AsString_ByEnum(vtReturnType)
    
    End If
    
    strCode = strCode & vbNewLine & vbNewLine & _
        Format.Format_AddTabToLines(strVBACode) & vbNewLine & vbNewLine & "End " & Coding.Get_Method_Type_AsString_ByEnum(vtType)

    
    Get_Code_Wrap_VBA_Method = strCode

End Function