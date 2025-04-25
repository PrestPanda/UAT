Option Compare Database
Option Explicit


Public Sub Listbox_Clear(strFormName As String, objListBox As Access.ListBox)
' Leert eine Access-Listbox unabhängig vom aktuellen RowSourceType über das Formularobjekt

    Dim objForm As Access.Form
    Set objForm = Forms(strFormName)

    With objForm.Controls(objListBox.Name)
        If .RowSourceType = "Table/Query" Or .RowSourceType = "Value List" Then
            .RowSource = ""
        End If
    End With

End Sub

Public Function ListBox_Get_Array( _
    strFormName As String, _
    objListBox As Object) As Variant()
' Gibt ein 2D-Array mit allen Einträgen der ListBox zurück (Zeilen x Spalten)

    Dim frm As Access.Form
    Dim intRows As Long
    Dim intCols As Long
    Dim intRow As Long
    Dim intCol As Long
    Dim varResult() As Variant

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        ListBox_Get_Array = Array()
        Exit Function
    End If

    Set frm = Forms(strFormName)

    With frm.Controls(objListBox.Name)
        intRows = .ListCount
        intCols = .ColumnCount

        If intRows = 0 Or intCols = 0 Then
            ListBox_Get_Array = Array()
            Exit Function
        End If

        ReDim varResult(1 To intRows, 1 To intCols)

        For intRow = 0 To intRows - 1
            For intCol = 0 To intCols - 1
                varResult(intRow + 1, intCol + 1) = Nz(.Column(intCol, intRow), "")
            Next intCol
        Next intRow
    End With

    ListBox_Get_Array = varResult

End Function
Public Function ListBox_Get_Array_Selected( _
    strFormName As String, _
    objListBox As Object) As String()
' Gibt die ausgewählten Einträge einer ListBox als Array zurück.

    Dim frm As Access.Form
    Dim i As Long, n As Long
    Dim arr() As String

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        ListBox_Get_Array_Selected = Split("") ' leeres Array
        Exit Function
    End If

    Set frm = Forms(strFormName)

    ReDim arr(0 To 0)
    n = -1

    With frm.Controls(objListBox.Name)
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                n = n + 1
                ReDim Preserve arr(0 To n)
                arr(n) = .ItemData(i)
            End If
        Next i
    End With

    If n = -1 Then
        ListBox_Get_Array_Selected = Split("") ' leeres Array
    Else
        ListBox_Get_Array_Selected = arr
    End If

End Function
Public Function ListBox_SetDefaultSettings(objListBox As Object)

    objListBox.RowSourceType = "Value List"
    objListBox.MultiSelect = fmMultiSelectSingle

End Function
Public Sub ListBox_Fill_FromArray(objListBox As MSForms.ListBox, varData As Variant)
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
Public Sub ListBox_RemoveValue( _
    strFormName As String, _
    strListBoxName As String, _
    strValue As String)

    ' Entfernt die gesamte Zeile aus der ListBox, wenn der übergebene Wert in einer der Spalten gefunden wird (nur bei RowSourceType = "Value List")

    Dim frm As Access.Form
    Dim strRowSource As String
    Dim arrRows() As String
    Dim strNewSource As String
    Dim i As Long
    Dim arrColumns() As String
    Dim intColumnIndex As Integer
    Dim intTotalColumns As Integer
    Dim blnDeleteRow As Boolean

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then Exit Sub

    Set frm = Forms(strFormName)

    With frm.Controls(strListBoxName)
        If .RowSourceType <> "Value List" Then
            MsgBox "ListBox '" & .Name & "' verwendet keinen Wertelistentyp!", vbExclamation
            Exit Sub
        End If

        strRowSource = Nz(.RowSource, "")
        If Len(strRowSource) = 0 Then Exit Sub

        arrRows = Split(strRowSource, ";")
        strNewSource = ""
        intTotalColumns = .ColumnCount

        ' Durchlaufe alle Zeilen
        For i = 0 To UBound(arrRows) Step intTotalColumns
            blnDeleteRow = False

            ' Prüfe jede Spalte innerhalb dieser Zeile
            Dim j As Long
            For j = 0 To intTotalColumns - 1
                If (i + j) <= UBound(arrRows) Then
                    If Trim(arrRows(i + j)) = strValue Then
                        blnDeleteRow = True
                        Exit For
                    End If
                End If
            Next j

            ' Wenn die Zeile nicht gelöscht werden soll, füge sie in den neuen RowSource ein
            If Not blnDeleteRow Then
                For j = 0 To intTotalColumns - 1
                    If (i + j) <= UBound(arrRows) Then
                        strNewSource = strNewSource & arrRows(i + j) & ";"
                    End If
                Next j
            End If
        Next i

        ' Entferne das letzte Semikolon
        If Right(strNewSource, 1) = ";" Then
            strNewSource = Left(strNewSource, Len(strNewSource) - 1)
        End If

        .RowSource = strNewSource
    End With

End Sub

Public Function ListBox_ContainsValue( _
    strFormName As String, _
    strListBoxName As String, _
    strValue As String) As Boolean
' Prüft, ob der übergebene Wert in einer beliebigen Spalte der ListBox enthalten ist

    Dim frm As Access.Form
    Dim intColumnIndex As Integer
    Dim intTotalColumns As Integer

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        ListBox_ContainsValue = False
        Exit Function
    End If

    Set frm = Forms(strFormName)

    intTotalColumns = frm.Controls(strListBoxName).ColumnCount

    For intColumnIndex = 0 To intTotalColumns - 1
        If ListBox_ContainsValue_InColumn(strFormName, strListBoxName, intColumnIndex, strValue) Then
            ListBox_ContainsValue = True
            Exit Function
        End If
    Next intColumnIndex

    ListBox_ContainsValue = False

End Function
Public Function ListBox_ContainsValue_InColumn( _
    strFormName As String, _
    strListBox_Name As String, _
    intColumnIndex As Integer, _
    strValue As String) As Boolean
' Prüft, ob der übergebene Wert in der angegebenen Spalte der ListBox enthalten ist

    Dim frm As Access.Form
    Dim lstListBox As Object
    Dim i As Long
    Dim strCurrent As String

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        ListBox_ContainsValue_InColumn = False
        Exit Function
    End If

    Set frm = Forms(strFormName)


    If intColumnIndex < 0 Or intColumnIndex > frm.Controls(strListBox_Name).ColumnCount Then
        MsgBox "Ungültige Spaltennummer (" & intColumnIndex & ") für ListBox '" & strListBox_Name & "'", vbCritical
        ListBox_ContainsValue_InColumn = False
        Exit Function
    End If
    
    For i = 0 To frm.Controls(strListBox_Name).ListCount - 1
        strCurrent = Nz(frm.Controls(strListBox_Name).Column(intColumnIndex, i), "")
        If strCurrent = strValue Then
            ListBox_ContainsValue_InColumn = True
            Exit Function
        End If
    Next i


    ListBox_ContainsValue_InColumn = False

End Function
Public Function ListBox_IsValueSelected( _
    strForm_Name As String, _
    strListBox_Name As String, _
    intColumnIndex As Integer, _
    strValue As String) As Boolean
' Prüft, ob ein bestimmter Eintrag in einer ListBox ausgewählt ist.
    
    Dim frm As Access.Form
    Dim lngIndex As Long

    If Not CurrentProject.AllForms(strForm_Name).IsLoaded Then
        ListBox_IsValueSelected = False
        Exit Function
    End If

    Set frm = Forms(strForm_Name)

    If Not ListBox_ContainsValue_InColumn(strForm_Name, strListBox_Name, intColumnIndex, strValue) Then
        ListBox_IsValueSelected = False
        Exit Function
    End If

    With frm.Controls(strListBox_Name)
        For lngIndex = 0 To .ListCount - 1
            If .Selected(lngIndex) Then
                If .Column(intColumnIndex, lngIndex) = strValue Then
                    ListBox_IsValueSelected = True
                    Exit Function
                End If
            End If
        Next lngIndex
    End With

    ListBox_IsValueSelected = False

End Function