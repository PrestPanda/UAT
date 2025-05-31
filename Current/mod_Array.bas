Option Compare Database
Option Explicit

Public Function Array_HasEntries(varArray As Variant) As Boolean

    ' Überprüft, ob ein übergebenes Array Einträge enthält

    On Error GoTo Fehler

    If IsArray(varArray) Then
        If Not IsEmpty(varArray) Then
            If (UBound(varArray) >= LBound(varArray)) Then
                Array_HasEntries = True
                Exit Function
            End If
        End If
    End If

    Array_HasEntries = False
    Exit Function

Fehler:
    Array_HasEntries = False

End Function

Public Function Array_GetStringTable(varArray As Variant) As String

    ' Gibt eine String-Tabelle aus einem 2D-Array zurück (Zeilen & Spalten, ohne Überschriften)

    Dim strOutput As String
    Dim lngRow As Long, lngCol As Long
    Dim strLine As String

    If Not IsArray(varArray) Then
        Array_GetStringTable = "<Kein gültiges Array>"
        Exit Function
    End If

    On Error GoTo Fehler
    For lngRow = LBound(varArray, 1) To UBound(varArray, 1)
        strLine = ""
        For lngCol = LBound(varArray, 2) To UBound(varArray, 2)
            If strLine <> "" Then strLine = strLine & vbTab
            strLine = strLine & Nz(varArray(lngRow, lngCol), "")
        Next lngCol
        strOutput = strOutput & strLine & vbCrLf
    Next lngRow

    Array_GetStringTable = strOutput
    Exit Function

Fehler:
    Array_GetStringTable = "<Fehler beim Verarbeiten des Arrays>"

    
End Function
Public Function Array_GetFromSQL(strSQL As String) As Variant()

 ' Führt eine SQL-Abfrage aus und gibt das Ergebnis als 2D-Array zurück.
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim lngRows As Long, lngCols As Long
    Dim varResults() As Variant
    Dim lngRow As Long, lngCol As Long

    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.EOF Then

        rs.Close
        Set rs = Nothing
        Set db = Nothing
        Array_GetFromSQL = varResults
        Exit Function
    End If

    lngCols = rs.Fields.Count
    rs.MoveLast
    lngRows = rs.RecordCount
    rs.MoveFirst

    ReDim varResults(1 To lngRows, 1 To lngCols)

    For lngRow = 1 To lngRows
        For lngCol = 1 To lngCols
            varResults(lngRow, lngCol) = rs.Fields(lngCol - 1).Value
        Next lngCol
        rs.MoveNext
    Next lngRow

    Array_GetFromSQL = varResults

    rs.Close
    Set rs = Nothing
    Set db = Nothing


End Function