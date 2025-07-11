Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
'Die Coding_VBA stellt Funktionalit?ten zur automatisierten Erstellung und Bearbeitung von VBA COde bereit


Private Coding_SQL As New clsCoding_SQL
Private Coding As New clsCoding
Private Coding_VBA_Write As New clsCoding_VBA_Write
Private Format As New clsFormat
Private Log As New clsLog


Public Function Get_Code_ClassModule(strClassName As String, _
    Properties() As Variant, _
    Methods() As Variant, _
    Classes() As Variant) As String
'Refactored

    Dim strCode As String
    Dim strTableName As String
    Dim strClassModuleName As String
    Dim lngCounter As Long
    
    strTableName = "tbl_" & strClassName
    strClassModuleName = "cls" & strClassName
  
    strCode = Get_Code_ClassModule_Class_References(Classes()) & vbNewLine
    strCode = strCode & vbNewLine & _
        Get_Code_ClassModule_PrivateVariables(Properties()) & vbNewLine
    strCode = strCode & vbNewLine & _
        Get_Code_ClassModule_Properties(Properties()) & vbNewLine
    strCode = strCode & vbNewLine & _
        Get_Code_Module_Methods(strTableName, Properties(), Methods()) & vbNewLine

    Get_Code_ClassModule = strCode

End Function
'-------------------------------------------- CODE - Module - PRIVATE VARIABLES -----------------------------------
Public Function Get_Code_ClassModule_Class_References(Classes() As Variant) As String
'Refactored

    Dim strCode As String
    Dim lngCounter As Long
    
    Log.WriteLine "Code f�r Verweise auf Klassen wird generiert."
    
    If Array_HasEntries(Classes()) = True Then
    
        For lngCounter = LBound(Classes()) To UBound(Classes())
        
            If strCode = "" Then
                strCode = Coding.Get_Class_Reference_DimensionStatement( _
                    Classes(lngCounter, 1)) & vbNewLine
            Else
            strCode = strCode & _
                Coding.Get_Class_Reference_DimensionStatement(Classes(lngCounter, 1)) & vbNewLine
        
            End If
        
        Next lngCounter
    
    End If
    
    strCode = strCode & vbNewLine
    
    Log.WriteLine "Code f�r Verweise auf Klassen wurde generiert."
    
    Get_Code_ClassModule_Class_References = strCode


End Function



Public Function Get_Code_ClassModule_PrivateVariables(Properties() As Variant) As String
'Refactored

    Dim lngCounter As Long

    Log.WriteLine "Code f�r private Variablen wird generiert."
    
    Dim strCode As String

    For lngCounter = LBound(Properties()) To UBound(Properties())

        If strCode = "" Then
        
            strCode = strCode & _
                Coding_VBA_Write.Get_Code_PrivateVariable_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2))
        Else
        
            strCode = strCode & vbNewLine & _
                Coding_VBA_Write.Get_Code_PrivateVariable_AsString_ByString _
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
             "Private", strName, strDataType)

End Function

'-------------------------------------------- CODE - Module - PROPERTIES -----------------------------------
Public Function Get_Code_ClassModule_Properties(Properties() As Variant) As String
'Refactored
    Dim strCode As String
    
    Log.WriteLine "Code f?r Properties wird generiert"
    
   For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If strCode = "" Then
            strCode = Get_Code_Module_Property_ByString( _
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
                            
    Get_Code_Module_Property_ByString = strCode
            
End Function
Public Function Get_Code_Module_Property_Get_ByString(strPropertyName As String, _
    strDataType As String, _
    strVisibility As String) As String
'Refactored

    Get_Code_Module_Property_Get_ByString = strVisibility & _
                            " Property Get " & strPropertyName & "() As " & Coding.Get_Variable_Translate_Datatype(strDataType) & _
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
                            Coding.Get_Variable_Translate_Datatype(strDataType) & ")" & _
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
    
    Log.WriteLine "Code f�r Methods wird generiert"
    
   For lngCounter = LBound(Methods()) To UBound(Methods())
    
        If IsNull(DLookup("Name", "tbl_Package_Method_Draft", _
            "Name = '" & Methods(lngCounter, 1) & "'")) = False Then
            
            'Code der Standardmethode in den Wrap einf?gen
            strCodeAddition = Get_Code_Module_Method_Draft( _
                strTableName, _
                DLookup("Name", "tbl_Package_Method_Draft", _
                "Name = '" & Methods(lngCounter, 1) & "'"), Properties())
            
        Else
        
            'Nur den Wrap erstellen
            strCodeAddition = Get_Code_Wrap_VBA_Method( _
                Coding.Get_Visibility_AsEnum_ByString(CStr(Methods(lngCounter, 2))), _
                Coding.Get_Method_Type_AsEnum_ByString(CStr(Methods(lngCounter, 3))), _
                CStr(Methods(lngCounter, 1)), "", "", vtNoType)
            
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
    strDraftName As String, _
    Properties() As Variant) As String
'Refactored

    Select Case DLookup("Package_FK", "tbl_Package_Method_Draft", "Name = '" & strDraftName & "'")
    
        Case DLookup("ID", "tbl_Package", "Name = '" & "DB" & "'")
            Get_Code_Module_Method_Draft = _
                Get_Code_Module_Method_Draft_DB(strTableName, strDraftName, Properties())
        
        Case DLookup("ID", "tbl_Package", "Name = '" & "Standard" & "'")
            Get_Code_Module_Method_Draft = _
                Get_Code_Module_Method_Draft_Standard(strTableName, strDraftName, Properties())
                
        Case DLookup("ID", "tbl_Package", "Name = '" & "Tracking" & "'")
        
        
        Case DLookup("ID", "tbl_Package", "Name = '" & "DEV" & "'")
        
    
    End Select

End Function
'-------------------------------------------- DRAFT Standard -----------------------------------
Public Function Get_Code_Module_Method_Draft_Standard(strTableName As String, _
    strDraftName As String, _
    Properties() As Variant) As String
'Refactored
    
    Dim strCode As String

    Select Case strDraftName
    
        Case "Reset"
            strCode = Get_Code_Module_Method_Draft_Standard_Reset(Properties())
                
'        Case "Validate_Add"
'            strCode = Get_Code_Module_Method_Draft_Standard_Validate_Add(strTableName, strDraftName, Properties())
        
        Case "Exists"
            strCode = Get_Code_Module_Method_Draft_Standard_Exists()
                
        Case "Delete"
            strCode = Get_Code_Module_Method_Draft_Standard_Delete()
                
        Case "Load"
            strCode = Get_Code_Module_Method_Draft_Standard_Load()
                
        Case "LoadByID"
            strCode = Get_Code_Module_Method_Draft_Standard_LoadByID()
                
        Case "Save"
            strCode = Get_Code_Module_Method_Draft_Standard_Save()
        
    
    End Select
    
    Get_Code_Module_Method_Draft_Standard = strCode

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

        strCode = strCode & "    " & Coding.Get_Variable_Name_AsString_ByString( _
            Properties(lngCounter, 1), _
            Properties(lngCounter, 2)) & " = " & strDefaultValue & vbCrLf
            
        
    Next lngCounter

    strCode = strCode & vbCrLf & "End Sub" & vbCrLf

    Get_Code_Module_Method_Draft_Standard_Reset = strCode

End Function
'Zun?chst On hold, da zun?chst die anderen Functions geschrieben werden m?ssen.
'Public Function Get_Code_Module_Method_Draft_Standard_Validate_Add(strTableName As String, _
'    strDraftName As String, _
'    Properties()) As String
'
'    Dim strCode As String
'
'    strCode = "Public Sub Validate_Add" & "()" & vbCrLf & vbCrLf
'
'    strCode = "if lngID = 0 then" & vbNewLine & vbNewLine & _
'        vbTab & "'To-Do: Hier Code f?r die Validierung einf?gen." & vbNewLine & _
'        vbTab & "Validate_Add = True" & vbNewLine & vbNewLine & _
'        "Else" & vbNewLine & vbNewLine & _
'        vbTab & "MsgBox"
'        vbTab & "Validate_Add = False" & vbNewLine & vbNewLine & _
'        "End if"
'
'
'    strCode = strCode & vbCrLf & "End Sub" & vbCrLf
'
'    Get_Code_Module_Method_Draft_Standard_Validate_Add = strCode
'
'End Function
Public Function Get_Code_Module_Method_Draft_Standard_Exists() As String

     Dim strCode As String

    strCode = _
        "If lngID = 0 Then" & vbCrLf & _
        vbTab & "Exists = False" & vbCrLf & _
        "Else" & vbCrLf & _
        vbTab & "If DB_Check = True Then" & vbCrLf & _
        vbTab & vbTab & "Exists = True" & vbCrLf & _
        vbTab & "Else" & vbCrLf & _
        vbTab & vbTab & "Exists = False" & vbCrLf & _
        vbTab & "End If" & vbCrLf & _
        "End If"

    Get_Code_Module_Method_Draft_Standard_Exists = _
        Get_Code_Wrap_VBA_Method(vPublic, mtFunction, "Exists", "", strCode, vtBoolean)

End Function
Public Function Get_Code_Module_Method_Draft_Standard_Delete() As String

    Dim strCode As String

    strCode = _
        "If Exists = True Then" & vbCrLf & _
        vbTab & "DB_Delete" & vbCrLf & _
        "End If"

    Get_Code_Module_Method_Draft_Standard_Delete = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "Delete", "", strCode, vtNoType)

End Function
Public Function Get_Code_Module_Method_Draft_Standard_Load() As String

    Dim strCode As String

    strCode = "DB_Load"

    Get_Code_Module_Method_Draft_Standard_Load = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "Load", "", strCode, vtNoType)

End Function
Public Function Get_Code_Module_Method_Draft_Standard_LoadByID() As String

    Dim strCode As String

    strCode = "lngID = lngID_2" & vbCrLf & _
        "Load" & vbCrLf

    Get_Code_Module_Method_Draft_Standard_LoadByID = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "LoadByID", "lngID_2 as long", strCode, vtNoType)

End Function
Public Function Get_Code_Module_Method_Draft_Standard_Save() As String

    Dim strCode As String

    strCode = _
        "If Exists = True Then" & vbCrLf & _
        vbTab & "DB_Update" & vbCrLf & _
        "Else" & vbCrLf & _
        vbTab & "DB_Insert" & vbCrLf & _
        "End If"

    Get_Code_Module_Method_Draft_Standard_Save = _
        Get_Code_Wrap_VBA_Method(vPublic, mtSub, "Save", "", strCode, vtNoType)

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

    Get_Code_Module_Method_Draft_DB_Activate = Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Activate", "", _
                                            "If DB_Check = true then " & vbNewLine & vbNewLine & _
                                            vbTab & "blnIsActive = True" & vbNewLine & _
                                            vbTab & "DB_Update" & vbNewLine & vbNewLine & _
                                            "End if", vtNoType)
End Function
Public Function Get_Code_Module_Method_Draft_DB_Check(strTableName As String) As String
'Refactored

    Dim strCode As String

    strCode = "if lngID = 0 then " & vbNewLine & vbNewLine & _
        vbTab & "DB_Check = False" & vbNewLine & vbNewLine & _
        "Else" & vbNewLine & vbNewLine & _
        vbTab & "If Isnull(dlookup(" & """ID""" & ",""" & strTableName & """," & """ID =""" & " & lngID)) = false then" & vbNewLine & _
        vbTab & vbTab & "DB_Check = true" & vbNewLine & _
        vbTab & "Else" & vbNewLine & _
        vbTab & vbTab & "DB_Check = false" & vbNewLine & _
        vbTab & "End If" & vbNewLine & vbNewLine & _
        "End If"

    Get_Code_Module_Method_Draft_DB_Check = _
        Get_Code_Wrap_VBA_Method(vPrivate, mtFunction, "DB_Check", "", strCode, vtBoolean)

End Function
Public Function Get_Code_Module_Method_Draft_DB_Deactivate(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Deactivate = Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Deactivate", "", _
                                            "If DB_Check = true then " & vbNewLine & vbNewLine & _
                                            vbTab & "blnIsActive = False" & vbNewLine & _
                                            vbTab & "DB_Update" & vbNewLine & vbNewLine & _
                                            "End if", vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Delete(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Delete = Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Delete", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Delete(strTableName)) & vbNewLine & vbNewLine & _
                                        "Reset", vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Insert(strTableName As String, Properties() As Variant) As String
'Refactored

    Get_Code_Module_Method_Draft_DB_Insert = Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Insert", "", _
                                        Get_Code_Module_Method_Draft_DB_Insert_SetValues(strTableName, Properties()) & vbNewLine & _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Insert(strTableName, _
                                        Properties())) & _
                                        Get_Code_Module_Method_DB_Insert_UpdateID(strTableName) _
                                        , vtNoType)

End Function
'Public Function Get_Code_Module_Method_DB_Insert_GetID(strTableName As String) As String
''Refactored
''Scheinbar nicht mehr n?tig wegen Get_Code_Module_Method_Draft_DB_Insert_SetValues
'
'    Get_Code_Module_Method_DB_Insert_GetID = vbNewLine & vbNewLine & _
'                                        "lngID = DMax(""" & "ID""" & ", """ & strTableName & """) + 1" & _
'                                        vbNewLine
'
'End Function
Public Function Get_Code_Module_Method_DB_Insert_UpdateID(strTableName As String) As String
'Refactored

    Get_Code_Module_Method_DB_Insert_UpdateID = vbNewLine & vbNewLine & _
                                        "lngID = DMax(""" & "ID""" & ", """ & strTableName & """)" & _
                                        vbNewLine

End Function
Public Function Get_Code_Module_Method_Draft_DB_Insert_SetValues(strTableName As String, _
    Properties() As Variant) As String
'New

    Dim lngCounter As Long
    Dim strCode As String

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        Select Case Properties(lngCounter, 1)
        
            Case "ID"
                strCode = "If IsNull(DMax(" & """ID""" & ", """ & strTableName & """)) = False then " & vbNewLine & _
                    vbTab & "lngID = DMax(" & """ID""" & ", """ & strTableName & """) + 1" & vbNewLine & _
                     "End if" & vbNewLine & vbNewLine

        
            Case "CreateTS", "LastUpdateTS"
                strCode = strCode & Coding.Get_Variable_Name_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2)) & _
                    " = now()" & vbNewLine
                
            Case "IsActive"
                strCode = strCode & Coding.Get_Variable_Name_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2)) & _
                    " = True " & vbNewLine
                    
        End Select
    
    
    Next lngCounter
    
    Get_Code_Module_Method_Draft_DB_Insert_SetValues = strCode

End Function
Public Function Get_Code_Module_Method_Draft_DB_Load(ByVal strTableName As String, _
     Properties() As Variant) As String
'Refactored

    Dim strCode As String
    Dim lngCounter As Long
    Dim strVarName As String

    strCode = "Dim rs As DAO.Recordset" & vbCrLf
    strCode = strCode & "Dim strSQL As String" & vbCrLf & vbCrLf
    strCode = strCode & "if db_check = true then " & vbNewLine & vbNewLine
    strCode = strCode & vbTab & "strSQL = ""SELECT * FROM " & strTableName & " WHERE ID = "" & lngID" & vbCrLf & vbCrLf
    strCode = strCode & vbTab & "Reset" & vbCrLf & vbCrLf
    strCode = strCode & vbTab & "Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)" & vbCrLf & vbCrLf
    strCode = strCode & vbTab & "If Not rs.EOF Then" & vbCrLf

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        strVarName = Coding.Get_Variable_Name_AsString_ByString( _
            Properties(lngCounter, 1), _
            Properties(lngCounter, 2))
            
        strCode = strCode & vbTab & "    " & strVarName & " = " & _
            Get_Code_Wrap_VBA_Load_Property( _
            CStr(Properties(lngCounter, 1)), _
            CStr(Properties(lngCounter, 2))) & vbCrLf

        
    Next lngCounter

    strCode = strCode & vbTab & "End If" & vbCrLf & vbCrLf
    strCode = strCode & vbTab & "rs.Close" & vbCrLf
    strCode = strCode & vbTab & "Set rs = Nothing" & vbCrLf & vbCrLf
    strCode = strCode & "end if" & vbCrLf


     Get_Code_Module_Method_Draft_DB_Load = _
        Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Load", "", strCode, vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Update(strTableName As String, _
    Properties() As Variant) As String
'Refactored
Dim strCode As String

    strCode = Get_Code_Module_Method_Draft_DB_Update_SetValues(Properties()) & vbNewLine & _
        Get_Code_Wrap_VBA_SQL_Statement( _
        Coding_SQL.Get_DB_Update(strTableName, Properties()))
        
    Get_Code_Module_Method_Draft_DB_Update = _
        Get_Code_Wrap_VBA_Method(vPrivate, mtSub, "DB_Update", "", strCode, vtNoType)


End Function
Public Function Get_Code_Module_Method_Draft_DB_Update_SetValues(Properties() As Variant) As String
'New

    Dim lngCounter As Long
    Dim strCode As String

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        Select Case Properties(lngCounter, 1)
        
            Case "LastUpdateTS"
                strCode = strCode & Coding.Get_Variable_Name_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2)) & _
                    " = now()" & vbNewLine
                    
        End Select
    
    
    Next lngCounter
    
    Get_Code_Module_Method_Draft_DB_Update_SetValues = strCode

End Function
'--------------------------------------------Code Wrap -----------------------------------
Public Function Get_Code_Wrap_VBA_SQL_Statement(strSQLStatement As String) As String

    Get_Code_Wrap_VBA_SQL_Statement = "docmd.setwarnings false" & vbNewLine & vbNewLine & _
                        "docmd.runsql " & strSQLStatement & vbNewLine & vbNewLine & _
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
Public Function Get_Code_Wrap_VBA_Load_Property( _
    strFieldName As String, _
    strDataType As String) As String

    ' Gibt den Code zur�ck, um einen Datenbankfeldwert sicher zu lesen (inkl. Nz(...) je nach Typ).
    
    Dim strResult As String

    Select Case strDataType
    
        Case "String", "Long String"
            strResult = "Nz(rs![" & strFieldName & "], """")"
            
        Case "Byte", "Integer", "Long", "Double"
            strResult = "Nz(rs![" & strFieldName & "], 0)"
            
        Case "Boolean"
            strResult = "Nz(rs![" & strFieldName & "], False)"
            
        Case "Date"
            strResult = "Nz(rs![" & strFieldName & "], #1/1/1900#)"
            
        Case Else
            ' Standard-Fallback auf einfachen Nz-Wrapper
            strResult = "Nz(rs![" & strFieldName & "])"
            
    End Select

    Get_Code_Wrap_VBA_Load_Property = strResult

End Function