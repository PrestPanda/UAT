Option Compare Database
Option Explicit

Public Function ListBox_ContainsValue( _
    strFormName As String, _
    objListBox As Object, _
    intColumnIndex As Integer, _
    strValue As String) As Boolean

    ' Prüft, ob der übergebene Wert in der angegebenen Spalte der ListBox enthalten ist

    
    Dim frm As Access.Form
    Dim i As Long
    Dim strCurrent As String

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then
        ListBox_ContainsValue = False
        Exit Function
    End If

    Set frm = Forms(strFormName)

    With frm.Controls(objListBox.Name)
        If intColumnIndex < 0 Or intColumnIndex > .ColumnCount - 1 Then
            MsgBox "Ungültige Spaltennummer (" & intColumnIndex & ") für ListBox '" & .Name & "'", vbCritical
            ListBox_ContainsValue = False
            Exit Function
        End If

        For i = 0 To .ListCount - 1
            strCurrent = Nz(.Column(intColumnIndex, i), "")
            If strCurrent = strValue Then
                ListBox_ContainsValue = True
                Exit Function
            End If
        Next i
    End With

    ListBox_ContainsValue = False

End Function
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
Public Sub Listbox_Clear(strFormName As String, objListBox As Access.Listbox)

    ' Leert eine Access-Listbox unabhängig vom aktuellen RowSourceType über das Formularobjekt

    
    Dim objForm As Access.Form
    Set objForm = Forms(strFormName)

    With objForm.Controls(objListBox.Name)
        If .RowSourceType = "Table/Query" Or .RowSourceType = "Value List" Then
            .RowSource = ""
        End If
    End With

End Sub
Public Function ListBox_Get_Array_Selected(objListBox As Listbox) As String()
    Dim i As Long, n As Long
    Dim arr() As String

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
        GetSelectedValues = Split("") ' leeres Array
    Else
        GetSelectedValues = arr
    End If
End Function