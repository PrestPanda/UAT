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
Public Function Format_Boolean(ByVal blnValue As Boolean) As String
    ' Formatiert einen Boolean-Wert für SQL-Insertion
    
    FormatSQLBoolean = IIf(blnValue, "True", "False")
    
End Function
Public Function Format_Date(ByVal dtmValue As Date) As String
    ' Formatiert ein Datum für SQL-Insertion
    
    FormatSQLDate = "#" & Format(dtmValue, "yyyy-mm-dd") & "#"
    
End Function
Public Function Format_Number(ByVal dblValue As Double) As String
    ' Formatiert eine Zahl für SQL-Insertion
    
    FormatSQLNumber = CStr(dblValue)
    
End Function
Public Function Format_String(ByVal strValue As String) As String
    ' Formatiert einen String-Wert für SQL-Insertion
    
    Format_String = "'" & Replace(strValue, "'", "") & "'"
    
End Function
Public Function Format_Value(ByVal varValue As Variant) As String
    ' Erkennt den Datentyp und ruft die entsprechende Formatierungsfunktion auf
    
    Select Case varType(varValue)
    
        Case vtString
            Format_Value = Format_String(varValue)
            
        Case vbDouble, vbSingle, vtlong, vbInteger, vbByte, vbCurrency
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
Private Function Get_DataType(PropertyType As enuCoding_Variable_Types) As String
    ' Wandelt den Variablentyp in den passenden SQL-Datentyp um
    Select Case PropertyType
    
        Case vtString
            Get_DataType = "TEXT(255)"
            
        Case vtStringLong
            Get_DataType = "LONGTEXT"
            
        Case vtlong
            Get_DataType = "INTEGER"
            
        Case vtDouble
            Get_DataType = "DOUBLE"
            
        Case vtDate
            Get_DataType = "DATETIME"
            
        Case vtBoolean
            Get_DataType = "YESNO"
            
        Case vtAuto
            Get_DataType = "AUTOINCREMENT"
            
        Case Else
            Get_DataType = "TEXT"
            
    End Select
End Function
Public Function Get_Code_Variable(varValue As Variant, varType As enuCoding_Variable_Types) As String
    ' Erkennt den Datentyp und erzeugt VBA-Code für SQL-Insertion
    
    Select Case varType
    
        Case vtString, vtStringLong
            Format_SQL_Code = "'"" & " & varValue & " & ""'"
        Case vtlong, vtDouble
            Format_SQL_Code = "" & varValue & ""
        Case vtDate
            Format_SQL_Code = "#" & varValue & "#"
        Case vtBoolean
            Format_SQL_Code = CInt(varType)
        Case vbNull
            Format_SQL_Code = "NULL"
        Case Else
            Err.Raise vbObjectError + 1000, "FormatSQLCode", "Unsupported data type"
    End Select
    
End Function

Public Function Get_DB_Activate(strTableName As String) As String
'Generates SQL Code for Activating a dataset from an object

    Get_DB_Activate = "UPDATE " & strTableName & " SET Active = true " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Deactivate(strTableName As String) As String
'Generates SQL Code for Deactivating a dataset from an object

    Get_DB_Deactivate = "UPDATE " & strTableName & " SET Active = false " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Delete(strTableName As String) As String
'Generates SQL Code for Deleting a dataset from an object

    Get_DB_Delete = "DELETE * FROM " & strTableName & " " & "WHERE ID =""" & "& lngID"

End Function
Public Function Get_DB_Insert(strTableName As String, strPropertyNames() As String, PropertyTypes() As enuCoding_Variable_Types) As String
'Generates VBA Code for Inserting a dataset into an object

    Dim i As Long
    Dim strColumns As String
    Dim strValues As String
    Dim strCode As String

    For i = LBound(strPropertyNames) To UBound(strPropertyNames)
        If strPropertyNames(i) <> "ID" Then
            strColumns = strColumns & strPropertyNames(i) & ", "
            strValues = strValues & Get_Code_Variable(Coding.Get_Variable_Name(strPropertyNames(i), PropertyTypes(i)), PropertyTypes(i)) & ", "
        End If
    Next i

    ' Entferne das letzte Komma und Leerzeichen
    strColumns = Left(strColumns, Len(strColumns) - 2)
    strValues = Left(strValues, Len(strValues) - 2)

    strCode = "INSERT INTO " & strTableName & " (" & strColumns & ") " & _
              "VALUES (" & strValues & ")"""

    Get_DB_Insert = strCode
End Function
Public Function Get_DB_Load(strTableName As String) As String
'Generates SQL Code for Loading a dataset from an object

    Get_DB_Load = "SELECT * FROM " & strTableName

End Function
Public Function Get_DB_Update(strTableName As String, strPropertyNames() As String, PropertyTypes() As enuCoding_Variable_Types) As String
'Generates SQL Code for Updating a dataset in an object

    Dim i As Long
    Dim strUpdates As String

    For i = LBound(strPropertyNames) To UBound(strPropertyNames)
        If strPropertyNames(i) <> "ID" Then
            strUpdates = strUpdates & strPropertyNames(i) & " = " & _
                Get_Code_Variable( _
                Coding.Get_Variable_Name(strPropertyNames(i), PropertyTypes(i)), PropertyTypes(i)) & ", "
        End If
    Next i

    ' Entferne das letzte Komma und Leerzeichen
    strUpdates = Left(strUpdates, Len(strUpdates) - 2)

    Get_DB_Update = "UPDATE " & strTableName & " SET " & strUpdates & " WHERE ID = """ & " &  lngID"

End Function
Public Function Get_DB_CreateTable(strTableName As String, strPropertyNames() As String, _
                                    PropertyTypes() As enuCoding_Variable_Types) As String
' Erstellt ein SQL-Statement zum Erstellen einer Tabelle
    
    Dim i As Long
    Dim strColumns As String
    Dim strSQL As String
    
    For i = LBound(strPropertyNames) To UBound(strPropertyNames)

            strColumns = strColumns & strPropertyNames(i) & " " & Get_DataType(PropertyTypes(i)) & ", "

        
    Next i
    
    
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