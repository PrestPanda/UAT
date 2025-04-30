Option Compare Database
Option Explicit

Private Function AccessListBox_Get_Object_ByName( _
    strFormName As String, _
    strListBoxName As String) As Access.ListBox

    ' Gibt das ListBox-Objekt eines angegebenen Formulars zurück

    Dim frm As Access.Form

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        Set AccessListBox_Get_Object_ByName = Nothing
        Exit Function
    End If

    Set frm = Forms(strFormName)


    Exit Function

Fehler:
    MsgBox "Die ListBox '" & strListBoxName & "' im Formular '" & strFormName & "' konnte nicht gefunden werden.", vbExclamation
    Set AccessListBox_Get_Object_ByName = Nothing

End Function
Public Sub Access_ListBox_Clear( _
    strFormName As String, _
    strListBoxName As String)

    ' Leert eine Access-Listbox unabhängig vom aktuellen RowSourceType

    Dim objListBox As Access.ListBox

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then Exit Sub

    If objListBox.RowSourceType = "Table/Query" Or objListBox.RowSourceType = "Value List" Then
        objListBox.RowSource = ""
    End If

End Sub
Public Function Access_ListBox_Get_Array( _
    strFormName As String, _
    strListBoxName As String) As Variant()

    ' Gibt ein 2D-Array mit allen Einträgen der ListBox zurück (Zeilen x Spalten)

    Dim objListBox As Access.ListBox
    Dim intRows As Long
    Dim intCols As Long
    Dim intRow As Long
    Dim intCol As Long
    Dim varResult() As Variant

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then
        Access_ListBox_Get_Array = Array()
        Exit Function
    End If

    intRows = objListBox.ListCount
    intCols = objListBox.ColumnCount

    If intRows = 0 Or intCols = 0 Then
        Access_ListBox_Get_Array = Array()
        Exit Function
    End If

    ReDim varResult(1 To intRows, 1 To intCols)

    For intRow = 0 To intRows - 1
        For intCol = 0 To intCols - 1
            varResult(intRow + 1, intCol + 1) = Nz(objListBox.Column(intCol, intRow), "")
        Next intCol
    Next intRow

    Access_ListBox_Get_Array = varResult

End Function
Public Function Access_ListBox_Get_Array_Selected( _
    strFormName As String, _
    strListBoxName As String) As String()

    ' Gibt die ausgewählten Einträge einer ListBox als Array zurück

    Dim objListBox As Access.ListBox
    Dim i As Long, n As Long
    Dim arr() As String

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then
        Access_ListBox_Get_Array_Selected = Split("") ' leeres Array
        Exit Function
    End If

    ReDim arr(0 To 0)
    n = -1

    For i = 0 To objListBox.ListCount - 1
        If objListBox.Selected(i) Then
            n = n + 1
            ReDim Preserve arr(0 To n)
            arr(n) = objListBox.ItemData(i)
        End If
    Next i

    If n = -1 Then
        Access_ListBox_Get_Array_Selected = Split("")
    Else
        Access_ListBox_Get_Array_Selected = arr
    End If

End Function
Public Function Access_ListBox_SetDefaultSettings( _
    strFormName As String, _
    strListBoxName As String)

    ' Setzt die Standardkonfiguration einer ListBox

    Dim objListBox As Access.ListBox

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then Exit Function

    objListBox.RowSourceType = "Value List"
    objListBox.MultiSelect = 1

End Function
Public Sub Access_ListBox_Fill_FromArray( _
    strFormName As String, _
    strListBoxName As String, _
    varData() As Variant)

    ' Füllt eine ListBox mit den Einträgen aus einem Array (1D oder 2D)
    
    Dim objListBox As Access.ListBox
    Dim intRow As Long
    Dim intCol As Long
    Dim intRows As Long
    Dim intCols As Long
    Dim strRowSource As String
    Dim strRow As String

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then Exit Sub

    Access_ListBox_Clear strFormName, strListBoxName

    On Error GoTo ExitSub

    intRows = UBound(varData, 1)
    intCols = UBound(varData, 2)

    ' 2D-Array erkannt ? RowSource zusammenbauen
    For intRow = 1 To intRows
        strRow = ""
        For intCol = 1 To intCols
            strRow = strRow & Nz(varData(intRow, intCol), "") & ";"
        Next intCol
        ' Semikolon am Ende entfernen
        If Right(strRow, 1) = ";" Then
            strRow = Left(strRow, Len(strRow) - 1)
        End If
        strRowSource = strRowSource & strRow & ";"
    Next intRow

    ' Gesamte RowSource setzen
    If Len(strRowSource) > 0 Then
        If Right(strRowSource, 1) = ";" Then
            strRowSource = Left(strRowSource, Len(strRowSource) - 1)
        End If
    End If

    objListBox.RowSourceType = "Value List"
    objListBox.RowSource = strRowSource

    Exit Sub

ExitSub:
    ' Falls Fehler (z.B. 1D-Array), mit AddItem arbeiten
    On Error Resume Next
    objListBox.RowSourceType = "Value List"
    objListBox.RowSource = ""

    For intRow = LBound(varData) To UBound(varData)
        objListBox.AddItem varData(intRow)
    Next intRow

End Sub

Public Sub Access_ListBox_RemoveValue( _
    strFormName As String, _
    strListBoxName As String, _
    strValue As String)

    ' Entfernt eine komplette Zeile aus der ListBox, wenn der übergebene Wert in einer Spalte gefunden wird

    Dim objListBox As Access.ListBox
    Dim strRowSource As String
    Dim arrRows() As String
    Dim strNewSource As String
    Dim i As Long
    Dim intColumnIndex As Long
    Dim intTotalColumns As Long
    Dim blnDeleteRow As Boolean
    Dim j As Long

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then Exit Sub

    If objListBox.RowSourceType <> "Value List" Then
        MsgBox "ListBox '" & objListBox.Name & "' verwendet keinen Wertelistentyp!", vbExclamation
        Exit Sub
    End If

    strRowSource = Nz(objListBox.RowSource, "")
    If Len(strRowSource) = 0 Then Exit Sub

    arrRows = Split(strRowSource, ";")
    strNewSource = ""
    intTotalColumns = objListBox.ColumnCount

    For i = 0 To UBound(arrRows) Step intTotalColumns
        blnDeleteRow = False

        For j = 0 To intTotalColumns - 1
            If (i + j) <= UBound(arrRows) Then
                If Trim(arrRows(i + j)) = strValue Then
                    blnDeleteRow = True
                    Exit For
                End If
            End If
        Next j

        If Not blnDeleteRow Then
            For j = 0 To intTotalColumns - 1
                If (i + j) <= UBound(arrRows) Then
                    strNewSource = strNewSource & arrRows(i + j) & ";"
                End If
            Next j
        End If
    Next i

    If Right(strNewSource, 1) = ";" Then
        strNewSource = Left(strNewSource, Len(strNewSource) - 1)
    End If

    objListBox.RowSource = strNewSource

End Sub
Public Function Access_ListBox_ContainsValue( _
    strFormName As String, _
    strListBoxName As String, _
    strValue As String) As Boolean

    ' Prüft, ob der übergebene Wert in einer beliebigen Spalte der ListBox enthalten ist

    Dim objListBox As Access.ListBox
    Dim intColumnIndex As Integer
    Dim intTotalColumns As Integer

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then
        Access_ListBox_ContainsValue = False
        Exit Function
    End If

    intTotalColumns = objListBox.ColumnCount

    For intColumnIndex = 0 To intTotalColumns - 1
        If Access_ListBox_ContainsValue_InColumn(strFormName, strListBoxName, intColumnIndex, strValue) Then
            Access_ListBox_ContainsValue = True
            Exit Function
        End If
    Next intColumnIndex

    Access_ListBox_ContainsValue = False

End Function
Public Function Access_ListBox_ContainsValue_InColumn( _
    strFormName As String, _
    strListBoxName As String, _
    intColumnIndex As Integer, _
    strValue As String) As Boolean

    ' Prüft, ob der übergebene Wert in einer bestimmten Spalte der ListBox enthalten ist

    Dim objListBox As Access.ListBox
    Dim i As Long
    Dim strCurrent As String

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then
        Access_ListBox_ContainsValue_InColumn = False
        Exit Function
    End If

    If intColumnIndex < 0 Or intColumnIndex > objListBox.ColumnCount - 1 Then
        MsgBox "Ungültige Spaltennummer (" & intColumnIndex & ") für ListBox '" & strListBoxName & "'", vbCritical
        Access_ListBox_ContainsValue_InColumn = False
        Exit Function
    End If

    For i = 0 To objListBox.ListCount - 1
        strCurrent = Nz(objListBox.Column(intColumnIndex, i), "")
        If strCurrent = strValue Then
            Access_ListBox_ContainsValue_InColumn = True
            Exit Function
        End If
    Next i

    Access_ListBox_ContainsValue_InColumn = False

End Function
Public Function Access_ListBox_IsValueSelected( _
    strFormName As String, _
    strListBoxName As String, _
    intColumnIndex As Integer, _
    strValue As String) As Boolean

    ' Prüft, ob ein bestimmter Eintrag in einer ListBox ausgewählt ist

    Dim objListBox As Access.ListBox
    Dim lngIndex As Long

    Set objListBox = AccessListBox_Get_Object_ByName(strFormName, strListBoxName)

    If objListBox Is Nothing Then
        Access_ListBox_IsValueSelected = False
        Exit Function
    End If

    If Not Access_ListBox_ContainsValue_InColumn(strFormName, strListBoxName, intColumnIndex, strValue) Then
        Access_ListBox_IsValueSelected = False
        Exit Function
    End If

    For lngIndex = 0 To objListBox.ListCount - 1
        If objListBox.Selected(lngIndex) Then
            If objListBox.Column(intColumnIndex, lngIndex) = strValue Then
                Access_ListBox_IsValueSelected = True
                Exit Function
            End If
        End If
    Next lngIndex

    Access_ListBox_IsValueSelected = False

End Function
Public Sub Access_Listbox_SelectedItem_Delete(strFormName As String, strControlName As String)

    Dim frmTarget As Form
    Dim ctlListbox As ListBox
    Dim intCols As Integer
    Dim intRows As Integer
    Dim intIndex As Integer
    Dim i As Long, j As Long
    Dim varData() As Variant
    Dim strRowSource As String
    Dim strNewRowSource As String
    Dim varTemp() As String

    Set frmTarget = Forms(strFormName)
    Set ctlListbox = frmTarget.Controls(strControlName)

    If ctlListbox.RowSourceType <> "Value List" Then
        MsgBox "Diese Funktion unterstützt nur ListBoxen mit RowSourceType = 'Value List'", vbExclamation
        Exit Sub
    End If

    If ctlListbox.ItemsSelected.Count = 0 Then
        MsgBox "Bitte wählen Sie eine Zeile aus.", vbExclamation
        Exit Sub
    End If

    intCols = ctlListbox.ColumnCount
    intRows = ctlListbox.ListCount
    intIndex = ctlListbox.ItemsSelected(0)

    ReDim varData(0 To intRows - 1, 0 To intCols - 1)
    For i = 0 To intRows - 1
        For j = 0 To intCols - 1
            varData(i, j) = ctlListbox.Column(j, i)
        Next j
    Next i

    strNewRowSource = ""
    For i = 0 To intRows - 1
        If i <> intIndex Then
            ReDim varTemp(0 To intCols - 1)
            For j = 0 To intCols - 1
                varTemp(j) = Nz(varData(i, j), "")
            Next j
            strNewRowSource = strNewRowSource & Join(varTemp, ";") & ";"
        End If
    Next i

    If Right(strNewRowSource, 1) = ";" Then
        strNewRowSource = Left(strNewRowSource, Len(strNewRowSource) - 1)
    End If

    ctlListbox.RowSource = strNewRowSource


    
End Sub
Public Sub Access_Listbox_SelectedItem_Move(strFormName As String, _
    strControlName As String, _
    enuDirection As enuDirection)

    ' Verschiebt die markierte Zeile in einer mehrspaltigen Access-ListBox (Value List) um eine Position.
    ' Unterstützt beliebig viele Spalten – funktioniert nur mit RowSourceType = "Value List".
    
    Dim frmTarget As Form
    Dim objListBox As ListBox
    Dim intCols As Integer
    Dim intRows As Integer
    Dim intIndex As Integer
    Dim varData() As Variant
    Dim i As Long, j As Long
    Dim strRowSource As String
    Dim varTemp() As String

    Set frmTarget = Forms(strFormName)
    Set objListBox = frmTarget.Controls(strControlName)

    If objListBox.RowSourceType <> "Value List" Then
        MsgBox "Diese Funktion unterstützt nur ListBoxen mit RowSourceType = 'Value List'", vbExclamation
        Exit Sub
    End If

    intCols = objListBox.ColumnCount
    intRows = objListBox.ListCount

    intIndex = -1
    For i = 0 To intRows - 1
        If objListBox.Selected(i) Then
            intIndex = i
            Exit For
        End If
    Next i

    If intIndex = -1 Then Exit Sub
    If enuDirection = Up And intIndex = 0 Then Exit Sub
    If enuDirection = Down And intIndex = intRows - 1 Then Exit Sub

    ReDim varData(0 To intRows - 1, 0 To intCols - 1)
    For i = 0 To intRows - 1
        For j = 0 To intCols - 1
            varData(i, j) = objListBox.Column(j, i)
        Next j
    Next i

    Dim rowA As Long, rowB As Long
    If enuDirection = Up Then
        rowA = intIndex
        rowB = intIndex - 1
    Else
        rowA = intIndex
        rowB = intIndex + 1
    End If

    Dim temp As Variant
    For j = 0 To intCols - 1
        temp = varData(rowA, j)
        varData(rowA, j) = varData(rowB, j)
        varData(rowB, j) = temp
    Next j

    strRowSource = ""
    For i = 0 To intRows - 1
        ReDim varTemp(0 To intCols - 1)
        For j = 0 To intCols - 1
            varTemp(j) = Nz(varData(i, j), "")
        Next j
        strRowSource = strRowSource & Join(varTemp, ";") & ";"
    Next i

    If Right(strRowSource, 1) = ";" Then
        strRowSource = Left(strRowSource, Len(strRowSource) - 1)
    End If

    objListBox.RowSource = strRowSource
    objListBox.Selected(rowB) = True

    
End Sub
Public Sub Access_ListBox_MovingButtons_UpdateActivation( _
    strFormName As String, _
    strListBoxName As String, _
    objButtonUp As Object, _
    objButtonDown As Object)

    ' Aktiviert oder deaktiviert die Buttons zum Verschieben eines Listbox-Eintrags je nach Auswahlposition
    ' Der Zugriff auf die Steuerelemente erfolgt direkt über Forms(strFormName).Controls("...").Enabled

    
    Dim intIndex As Long
    Dim intCount As Long

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then Exit Sub

    intCount = Forms(strFormName).Controls(strListBoxName).ListCount
    intIndex = Forms(strFormName).Controls(strListBoxName).ListIndex + 1

    ' Wenn nichts ausgewählt oder Liste leer ? Buttons deaktivieren
    If intCount = 0 Or Forms(strFormName).Controls(strListBoxName).ListIndex = -1 Then
        Forms(strFormName).Controls(objButtonUp.Name).Enabled = False
        Forms(strFormName).Controls(objButtonDown.Name).Enabled = False
        Exit Sub
    End If

    ' Standardmäßig beide aktivieren
    Forms(strFormName).Controls(objButtonUp.Name).Enabled = True
    Forms(strFormName).Controls(objButtonDown.Name).Enabled = True

    ' Randposition prüfen
    If intIndex = 1 Then Forms(strFormName).Controls(objButtonUp.Name).Enabled = False
    If intIndex = intCount Then Forms(strFormName).Controls(objButtonDown.Name).Enabled = False

End Sub