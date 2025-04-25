Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Dim Coding As New clsCoding
Dim Coding_SQL As New clsCoding_SQL

Public Enum enuCoding_SQL_DataTypes
    
    sqldtDate
    sqldtString

End Enum
'-------------------------------------------- FORMAT -----------------------------------
Public Function Format_Value(ByVal varValue As Variant) As String
'Refactored
    ' Erkennt den Datentyp und ruft die entsprechende Formatierungsfunktion auf
    
    Select Case varType(varValue)
    
        Case vtString
            Format_Value = Format_String(varValue)
            
        Case vbDouble, vbSingle, vtLong, vbInteger, vbByte, vbCurrency
            Format_Value = Format_Number(varValue)
            
        Case vbDate
            Format_Value = Format_Date(varValue)
            
        Case vbBoolean
            Format_Value = Format_Boolean(varValue)
            
        Case vbNull
            Format_Value = "NULL"
            
        Case Else
            Err.Raise vbObjectError + 1000, "FormatSQLValue", "Unsupported data type"
            
    End Select
    
End Function
Public Function Format_Boolean(ByVal blnValue As Boolean) As String
'Refactored
    ' Formatiert einen Boolean-Wert für SQL-Insertion
    
    FormatSQLBoolean = IIf(blnValue, "True", "False")
    
End Function
Public Function Format_Date(ByVal dtmValue As Date) As String
'Refactored
    ' Formatiert ein Datum für SQL-Insertion
    
    FormatSQLDate = "#" & Format(dtmValue, "yyyy-mm-dd") & "#"
    
End Function
Public Function Format_Number(ByVal dblValue As Double) As String
'Refactored
    ' Formatiert eine Zahl für SQL-Insertion
    
    FormatSQLNumber = CStr(dblValue)
    
End Function
Public Function Format_String(ByVal strValue As String) As String
'Refactored
    ' Formatiert einen String-Wert für SQL-Insertion
    
    Format_String = "'" & Replace(strValue, "'", "") & "'"
    
End Function


'-------------------------------------------- CODE - METHODS - DB_ -----------------------------------

Public Function Get_DB_CreateTable(strTableName As String, Properties() As Variant) As String
'In refactoring
' Erstellt ein SQL-Statement zum Erstellen einer Tabelle
    
    Dim lngCounter As Long
    Dim strColumns As String
    Dim strSQL As String
    
    For lngCounter = LBound(Properties()) To UBound(Properties())

            strColumns = strColumns & _
                Properties(lngCounter, 1) & " " & _
                Get_DataType_AsString_ByString(CStr(Properties(lngCounter, 2))) & ", "

        
    Next lngCounter
    
    
    ' Entferne das letzte Komma und Leerzeichen
    strColumns = Left(strColumns, Len(strColumns) - 2)
    
    ' SQL-Statement erstellen
    
    If InStr(strColumns, "ID") <> 0 Then
        strSQL = "CREATE TABLE " & strTableName & " (" & strColumns & "," & _
         vbNewLine & "CONSTRAINT pkID PRIMARY KEY (ID));"
    Else
        strSQL = "CREATE TABLE " & strTableName & " (" & strColumns & ");"
    End If
    
    
    ' Gibt das SQL-Statement zurück

    Get_DB_CreateTable = strSQL
    
End Function
Public Function Get_DB_Activate(strTableName As String) As String
'Refactored
'Generates SQL Code for Activating a dataset from an object

    Get_DB_Activate = "UPDATE " & strTableName & " SET Active = true " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Deactivate(strTableName As String) As String
'Refactored
'Generates SQL Code for Deactivating a dataset from an object

    Get_DB_Deactivate = "UPDATE " & strTableName & " SET Active = false " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Delete(strTableName As String) As String
'Refactored
'Generates SQL Code for Deleting a dataset from an object

    Get_DB_Delete = "DELETE * FROM " & strTableName & " " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Insert(strTableName As String, Properties() As Variant) As String
'In Refactoring
'Generates VBA Code for Inserting a dataset into an object

    Dim lngCounter As Long
    Dim strColumns As String
    Dim strValues As String
    Dim strCode As String

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If Properties(lngCounter, 1) <> "ID" Then
        
            strColumns = strColumns & Properties(lngCounter, 1) & ", "
            
            strValues = strValues & _
                Coding_SQL.Get_Code_Variable_AsString_ByString( _
                    Properties(lngCounter, 1), _
                    Properties(lngCounter, 2)) & ", "
        End If
        
    Next lngCounter

    ' Entferne das letzte Komma und Leerzeichen
    strColumns = Left(strColumns, Len(strColumns) - 2)
    strValues = Left(strValues, Len(strValues) - 2)

    strCode = "INSERT INTO " & strTableName & vbNewLine & _
                " (" & strColumns & ") " & vbNewLine & _
                "VALUES (" & strValues & ")"""

    Get_DB_Insert = strCode
    
End Function
Public Function Get_DB_Load(strTableName As String) As String
'Generates SQL Code for Loading a dataset from an object

    Get_DB_Load = "SELECT * FROM " & strTableName

End Function
Public Function Get_DB_Update(strTableName As String, Properties() As Variant) As String
'Refactored
'Generates SQL Code for Updating a dataset in an object

    Dim lngCounter As Long
    Dim strUpdates As String

    For lngCounter = LBound(Properties()) To UBound(Properties())
    
        If Properties(lngCounter, 1) <> "ID" Then
            strUpdates = strUpdates & Properties(lngCounter, 1) & " = " & _
                Coding.Get_Variable_Name_AsString_ByString( _
                    CStr(Properties(lngCounter, 1)), _
                    CStr(Properties(lngCounter, 2))) & ", "
        End If
        
    Next lngCounter

    ' Entferne das letzte Komma und Leerzeichen
    strUpdates = Left(strUpdates, Len(strUpdates) - 2)

    Get_DB_Update = "UPDATE " & strTableName & " SET " & strUpdates & " WHERE ID = """ & " &  lngID"

End Function

Private Function Get_DataType_AsString_ByString(strDataType As String) As String
'Refactored
    ' Wandelt den Variablentyp in den passenden SQL-Datentyp um
    Select Case strDataType
    
        Case "String"
            Get_DataType_AsString_ByString = "TEXT(255)"
            
        Case "StringLong"
            Get_DataType_AsString_ByString = "LONGTEXT"
            
        Case "Long"
            Get_DataType_AsString_ByString = "INTEGER"
            
        Case "Double"
            Get_DataType_AsString_ByString = "DOUBLE"
            
        Case "Date"
            Get_DataType_AsString_ByString = "DATETIME"
            
        Case "Boolean"
            Get_DataType_AsString_ByString = "YESNO"
            
        Case "Auto"
            Get_DataType_AsString_ByString = "AUTOINCREMENT"
            
        Case Else
            Get_DataType_AsString_ByString = "TEXT"
            
    End Select
End Function
Private Function Get_DataType_AsString_ByEnum(DataType As enuCoding_DataTypes) As String
'Refactored
    ' Wandelt den Variablentyp in den passenden SQL-Datentyp um
    
    Select Case PropertyType
    
        Case vtString
            Get_DataType_AsString_ByEnum = "TEXT(255)"
            
        Case vtStringLong
            Get_DataType_AsString_ByEnum = "LONGTEXT"
            
        Case vtLong
            Get_DataType_AsString_ByEnum = "INTEGER"
            
        Case vtDouble
            Get_DataType_AsString_ByEnum = "DOUBLE"
            
        Case vtDate
            Get_DataType_AsString_ByEnum = "DATETIME"
            
        Case vtBoolean
            Get_DataType_AsString_ByEnum = "YESNO"
            
        Case vtAuto
            Get_DataType_AsString_ByEnum = "AUTOINCREMENT"
            
        Case Else
            Get_DataType_AsString_ByEnum = "TEXT"
            
    End Select
    
End Function
Public Function Get_Code_Variable_AsString_ByString(ByVal strVariableName As String, ByVal strDataType As String) As String
'Refactored
    ' Erkennt den Datentyp und erzeugt VBA-Code für SQL-Insertion
    
    Select Case Coding.Get_Variable_Type_AsEnum_ByString(strDataType)
    
        Case vtString, vtStringLong
            Get_Code_Variable_AsString_ByString = "'"" & " & strVariableName & " & ""'"
            
        Case vtLong, vtDouble
            Get_Code_Variable_AsString_ByString = "" & strVariableName & ""
            
        Case vtDate
            Get_Code_Variable_AsString_ByString = "#" & strVariableName & "#"
            
        Case vtBoolean
            Get_Code_Variable_AsString_ByString = CInt(Coding.Get_Variable_Type_AsEnum_ByString(strDataType))
            
        Case vbNull
            Get_Code_Variable_AsString_ByString = "NULL"
            
    End Select
    
End Function