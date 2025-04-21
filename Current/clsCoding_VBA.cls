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

'--------------------------------------------Code Variable -----------------------------------

Public Function Get_Code_ClassModule(strClassName As String, Properties() As Variant, Methods() As Variant) As String

    Dim strCode As String
    Dim strTableName As String
    Dim strClassModuleName As String
    Dim lngCounter As Long
    
    strTableName = "tbl_" & strClassName
    strClassModuleName = "cls" & strClassName
  
    
    
    'Private Variables
    Log.WriteLine "Code für private Variablen wird generiert."
    
    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If strCode = "" Then
            strCode = strCode & _
                Coding_VBA.Get_Code_PrivateVariable( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2))
        Else
            strCode = strCode & vbNewLine & _
                Coding_VBA.Get_Code_PrivateVariable _
                    (Properties(lngCounter, 1), _
                    Properties(lngCounter, 2))
        End If
   
    Next
       
    Log.WriteLine "Code generiert."
    
    strCode = strCode & vbNewLine
    
    'Properties
    Log.WriteLine "Code für Properties wird generiert"
    
    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If strCode = "" Then
            strCode = strCode & _
                Get_Code_Module_Property( _
                strPropertyName(lngCounter), _
                strPropertyType(lngCounter), vPublic)
        Else
            strCode = strCode & vbNewLine & Get_Code_Module_Property(strPropertyName(lngCounter), strPropertyType(lngCounter), vPublic)
        End If
      
    
    Next
    
    strCode = strCode & vbNewLine
    
    
    'Standard_Methods

    strCode = strCode & Get_Code_Module_Methods_Standard(strPropertyName(), strPropertyType()) & vbNewLine
    
   'DB_Methods
    
    If blnCreateDBMethods = True Then

        strCode = strCode & Get_Code_Module_Methods_DB("tbl_" & strClassName, strPropertyName(), strPropertyType())
    End If
    
    
    
    'Erstelle Property
Get_Code_Module = strCode

End Function


'-------------------------------------------- CODE - Module - PROPERTIES -----------------------------------


'-------------------------------------------- CODE - Module - METHODS - Standard -----------------------------------
Public Function Get_Code_Module_Methods_Standard(PropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String

        Get_Code_Module_Methods_Standard = Get_Code_Module_Methods_Standard_Clear(PropertyNames(), PropertyTypes()) & vbNewLine & _
                                            Get_Code_Module_Methods_Standard_FullObjectText_Debug

End Function

Public Function Get_Code_Module_Methods_Standard_Clear(PropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String
    ' Gibt einen String zurück, der Code enthält, um alle Properties des angegebenen
    ' Klassenmoduls auf einen Defaultwert zu setzen. Die Properties und deren Typen werden
    ' über die Parameter PropertyNames und PropertyTypes übergeben.

    Dim strCode As String
    Dim lngIndex As Long
    Dim strDefaultValue As String

    ' Erstelle den Methodenkopf mit dem übergebenen Tabellennamen
    strCode = "Public Sub Clear" & strTableName & "()" & vbCrLf & vbCrLf

    ' Durchlaufe alle übergebenen Properties und weise ihnen den Defaultwert zu
    For lngIndex = LBound(PropertyNames) To UBound(PropertyNames)
        Select Case PropertyTypes(lngIndex)
        
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

        strCode = strCode & "    Me." & PropertyNames(lngIndex) & " = " & strDefaultValue & vbCrLf
    Next lngIndex

    strCode = strCode & vbCrLf & "End Sub" & vbCrLf

    Get_Code_Module_Methods_Standard_Clear = strCode

End Function


Public Function Get_Code_Module_Methods_Standard_FullObjectText_Debug()





End Function

'-------------------------------------------- CODE - Module - METHODS - DB -----------------------------------
Public Function Get_Code_Module_Methods_DB(strTableName As String, PropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String

    
    Get_Code_Module_Methods_DB = Get_Code_Module_Method_DB_Activate(strTableName) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Check(strTableName) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Deactivate(strTableName) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Delete(strTableName) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Insert(strTableName, PropertyNames(), PropertyTypes()) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Load(strTableName, PropertyNames(), PropertyTypes()) & vbNewLine & _
                                    Get_Code_Module_Method_DB_Update(strTableName, PropertyNames(), PropertyTypes())
                                    
End Function
Public Function Get_Code_Module_Method_DB_Activate(strTableName As String) As String


    Get_Code_Module_Method_DB_Activate = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Activate", "", _
                                            Get_Code_Wrap_VBA_SQL_Statement( _
                                            Coding_SQL.Get_DB_Activate(strTableName)) _
                                            , vtNoType)
    
End Function
Public Function Get_Code_Module_Method_DB_Check(strTableName As String) As String

    Dim strCode As String
    
    strCode = "If Isnull(dlookup(" & """ID""" & ",""" & strTableName & """," & """ID =""" & " & lngID)) = false then" & vbNewLine & _
        vbTab & "DB_Check = true" & vbNewLine & _
        "Else" & vbNewLine & _
        vbTab & "DB_Check = false" & vbNewLine & _
        "End If"

    Get_Code_Module_Method_DB_Check = Get_Code_Wrap_VBA_Method(vPublic, mtFunction, "DB_Check", "", strCode, vtBoolean)

        
End Function
Public Function Get_Code_Module_Method_DB_Deactivate(strTableName As String) As String

    Get_Code_Module_Method_DB_Deactivate = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Deactivate", "", _
                                            Get_Code_Wrap_VBA_SQL_Statement( _
                                            Coding_SQL.Get_DB_Deactivate(strTableName)) _
                                            , vtNoType)
            
End Function
Public Function Get_Code_Module_Method_DB_Delete(strTableName As String) As String

    Get_Code_Module_Method_DB_Delete = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Delete", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Delete(strTableName)) _
                                        , vtNoType)

End Function
Public Function Get_Code_Module_Method_DB_Insert(strTableName As String, strPropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String

    Get_Code_Module_Method_DB_Insert = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Insert", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                        Coding_SQL.Get_DB_Insert(strTableName, strPropertyNames(), PropertyTypes())) & _
                                        Get_Code_Module_Method_DB_Insert_UpdateID(strTableName) _
                                        , vtNoType)

End Function
Public Function Get_Code_Module_Method_DB_Insert_UpdateID(strTableName As String) As String

    Get_Code_Module_Method_DB_Insert_UpdateID = vbNewLine & vbNewLine & _
                                        "lngID = DMax(""" & "ID""" & ", """ & strTableName & """)" & _
                                        vbNewLine

End Function
Public Function Get_Code_Module_Method_DB_Load(strTableName As String, PropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String

    Dim strCode As String
    Dim intIndex As Integer
        Dim strVarName As String
    Dim strVarType As String
    
    strCode = "Dim rs As DAO.Recordset" & vbCrLf
    strCode = strCode & "Dim strSQL As String" & vbCrLf
    strCode = strCode & "strSQL = ""SELECT * FROM " & strTableName & " WHERE ID = "" & lngID" & vbCrLf
    strCode = strCode & "Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenDynaset)" & vbCrLf
    strCode = strCode & "If Not rs.EOF Then" & vbCrLf

    For intIndex = LBound(PropertyNames) To UBound(PropertyNames)
        strVarName = Coding.Get_Variable_Prefix(PropertyTypes(intIndex)) & PropertyNames(intIndex)  ' Präfix für private Variablen
        strVarType = Coding.Get_Variable_Type(PropertyTypes(intIndex))
        strCode = strCode & "    " & strVarName & " = rs![" & PropertyNames(intIndex) & "]" & vbCrLf
    Next intIndex

    strCode = strCode & "End If" & vbCrLf
    strCode = strCode & "rs.Close" & vbCrLf
    strCode = strCode & "Set rs = Nothing" & vbCrLf

    
     Get_Code_Module_Method_DB_Load = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Load", "", strCode, vtNoType)

End Function
Public Function Get_Code_Module_Method_DB_Update(strTableName As String, strPropertyNames() As String, PropertyTypes() As enuCoding_DataTypes) As String

    Get_Code_Module_Method_DB_Update = Get_Code_Wrap_VBA_Method(vPublic, mtSub, "DB_Update", "", _
                                        Get_Code_Wrap_VBA_SQL_Statement( _
                                            Coding_SQL.Get_DB_Update(strTableName, strPropertyNames(), PropertyTypes())) _
                                    , vtNoType)

End Function
'-------------------------------------------- CODE - Module - SORT INPUT -----------------------------------


'Private Function Get_Code_Module_Sort_Input(strPropertyNames() As String, PropertyType As enuCoding_DataTypes)
''Eventuell später benötigt, aber jetzt noch nicht
''Sortiert die entgegengenommenen PropertyNames und Types
'
'End Function

'-------------------------------------------- CODE - Module - PRIVATE VARIABLES-----------------------------------
Public Function Get_Code_PrivateVariable(strName As String, strDataType As String) As String
'Refactored

    Get_Code_Module_PrivateVariable = _
        Coding.Get_Variable_DimensionStatement( _
            vPrivate, strName, strDataType)

End Function
Public Function Get_Code_Module_PrivateVariables_Standard() As String

    Get_Code_Module_PrivateVariables = Coding.Get_Variable_DimensionStatement(vPrivate, "ID", vtAuto)

End Function
Public Function Get_Code_Module_Property(strPropertyName As String, PropertyType As enuCoding_DataTypes, Visibility As enuCoding_Visibility) As String

    Dim strCode As String

    strCode = Coding.Get_Visibility(Visibility) & _
                            " Property Get " & strPropertyName & "() As " & Coding.Get_Variable_Type(PropertyType) & _
                             vbNewLine & vbNewLine & _
                             vbTab & strPropertyName & " = " & Coding.Get_Variable_Prefix(PropertyType) & strPropertyName & _
                             vbNewLine & vbNewLine & _
                            "End Property"
                            

        
    strCode = strCode & vbNewLine & Coding.Get_Visibility(Visibility) & _
                            " Property Let " & strPropertyName & "(" & _
                            Coding.Get_Variable_Prefix(PropertyType) & strPropertyName & "_1 as " & _
                            Coding.Get_Variable_Type(PropertyType) & ")" & _
                             vbNewLine & vbNewLine & _
                             vbTab & Coding.Get_Variable_Prefix(PropertyType) & strPropertyName & " = " & _
                             Coding.Get_Variable_Prefix(PropertyType) & strPropertyName & "_1" & _
                             vbNewLine & vbNewLine & _
                            "End Property"
                            
    Get_Code_Module_Property = strCode
            
End Function

'--------------------------------------------Code Wrap -----------------------------------
Public Function Get_Code_Wrap_VBA_SQL_Statement(strSQLStatement As String) As String

    Get_Code_Wrap_VBA_SQL_Statement = "docmd.setwarnings false" & vbNewLine & vbNewLine & _
                        "docmd.runsql """ & strSQLStatement & vbNewLine & vbNewLine & _
                        "docmd.setwarnings true"

End Function
'Public Function Get_Code_Wrap_VBA_Property(Visibility as enuCoding_Visibility, ) As String
'
'
'
'End Function
Public Function Get_Code_Wrap_VBA_Method(Visability As enuCoding_Visibility, vtType As enuCoding_Method_Types, _
    strMethodName As String, strParameters As String, strVBACode As String, vtReturnType As enuCoding_DataTypes)
    
    Dim strCode As String
   

    If vtReturnType = vtNoType Then
        strCode = Coding.Get_Visibility(Visability) & " " & Coding.Get_Method_Type(vtType) & " " & _
            strMethodName & "(" & strParameters & ")"
    Else
        strCode = Coding.Get_Visibility(Visability) & " " & Coding.Get_Method_Type(vtType) & " " & _
            strMethodName & "(" & strParameters & ") as " & Coding.Get_Variable_Type(vtReturnType)
    
    End If
    
    strCode = strCode & vbNewLine & vbNewLine & _
        Format.Format_AddTabToLines(strVBACode) & vbNewLine & vbNewLine & "End " & Coding.Get_Method_Type(vtType)

    
    Get_Code_Wrap_VBA_Method = strCode

End Function



'PROPERTIES