Option Compare Database
Option Explicit

Public Sub FillListBoxFromArray(objListBox As MSForms.Listbox, varData As Variant)

    ' Füllt eine Listbox mit den Einträgen aus einem Array
    ' Unterstützt sowohl 1D- als auch 2D-Arrays

    
    Dim intRow As Long
    Dim intCol As Long
    Dim intRows As Long
    Dim intCols As Long
    Dim varRow() As Variant

    objListBox.Clear

    On Error GoTo ExitSub
    intRows = UBound(varData, 1)
    intCols = UBound(varData, 2)
    ' ? 2D-Array erkannt

    For intRow = 1 To intRows
        ReDim varRow(0 To intCols - 1)
        For intCol = 1 To intCols
            varRow(intCol - 1) = varData(intRow, intCol)
        Next intCol
        objListBox.AddItem varRow(0)
        For intCol = 1 To intCols - 1
            objListBox.List(objListBox.ListCount - 1, intCol) = varRow(intCol)
        Next intCol
    Next intRow
    Exit Sub

ExitSub:
    ' Falls 1D-Array, wird hier weitergemacht
    On Error Resume Next
    objListBox.Clear
    For intRow = LBound(varData) To UBound(varData)
        objListBox.AddItem varData(intRow)
    Next intRow

End Sub
Public Function Get_Array_FromQuery(strQueryName As String) As Variant

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
            varResult(intRow, intCol) = rs.Fields(intCol - 1).value
        Next intCol
        rs.MoveNext
    Next intRow

    Get_Array_FromQuery = varResult

End Function
Public Function Get_Listbox_Selected(objListBox As Listbox) As Variant

 ' Gibt alle ausgewählten Einträge einer Listbox als Array zurück

    
    Dim lngIndex As Long
    Dim intSelectedCount As Integer
    Dim varSelected() As Variant

    ' Zähle ausgewählte Elemente
    For lngIndex = 0 To objListBox.ListCount - 1
        If objListBox.Selected(lngIndex) Then
            intSelectedCount = intSelectedCount + 1
        End If
    Next lngIndex

    If intSelectedCount = 0 Then
        Get_Listbox_Selected = Array()
        Exit Function
    End If

    ReDim varSelected(0 To intSelectedCount - 1)

    intSelectedCount = 0
    For lngIndex = 0 To objListBox.ListCount - 1
        If objListBox.Selected(lngIndex) Then
            varSelected(intSelectedCount) = objListBox.ItemData(lngIndex)
            intSelectedCount = intSelectedCount + 1
        End If
    Next lngIndex

    Get_Listbox_Selected = varSelected
    

End Function
Public Sub ClearListBoxEntries( _
    strFormName As String, _
    objListBox As Object)

    ' Löscht alle Einträge einer ListBox mit RowSourceType "Value List" auf dem angegebenen Formular

    
    If Not CurrentProject.AllForms(strFormName).IsLoaded Then Exit Sub

    With Forms(strFormName).Controls(objListBox.Name)
        If .RowSourceType = "Value List" Then .RowSource = ""
    End With

End Sub

'ListBox
'Herkunftstyp: Wertliste
'Mehrfachauswahl: Einzeln