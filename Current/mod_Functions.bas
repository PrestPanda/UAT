Option Compare Database
Option Explicit


Public Function Get_Array_FromQuery(strQueryName As String) As Variant()

    ' Führt eine gespeicherte Access-Query aus und gibt das Ergebnis als 2D-Array zurück

    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varResult() As Variant
    Dim intRowCount As Integer
    Dim intFieldCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Set db = CurrentDb
    Set rs = db.OpenRecordset(strQueryName, dbOpenSnapshot)
    

    If rs.EOF Then
        Get_Array_FromQuery = Array() ' leeres Array
        Exit Function
    End If

    intFieldCount = rs.Fields.Count
    rs.MoveLast
    intRowCount = rs.RecordCount
    rs.MoveFirst

    ReDim varResult(1 To intRowCount, 1 To intFieldCount)

    For intRow = 1 To intRowCount
        For intCol = 1 To intFieldCount
            varResult(intRow, intCol) = rs.Fields(intCol - 1).Value
        Next intCol
        rs.MoveNext
    Next intRow

    Get_Array_FromQuery = varResult

End Function