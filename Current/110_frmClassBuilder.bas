VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_110_frmClassBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum enuDirection
    Up
    Down
End Enum

Private Sub cmdAddMethod_Click()

    Dim strVisability As String
    Dim strType As String

    If optMethodAddPrivate = False And optMethodAddPublic = False Then
        MsgBox "Bitte legen Sie die Sichbarkeit der Methode fest."
        Exit Sub
    End If
    
    If optMethodAddTypeFunction = False And optMethodAddTypeSub = False Then
        MsgBox "Bitte legen Sie den Typ der Methode fest."
    End If
    
    If txtAddMethodName <> "" Then
    
        If optMethodAddPrivate = True Then strVisability = "Private"
        If optMethodAddPublic = True Then strVisability = "Public"
        
        If optMethodAddTypeFunction = True Then strType = "Function"
        If optMethodAddTypeSub = True Then strType = "Sub"

        
         lstPreviewMethods.AddItem txtAddMethodName.value & ";" & _
            strType & ";" & strVisability
            
        txtAddMethodName.value = ""
        ApplyDefaultSettings
    Else
        MsgBox "Es wurde kein Name für die Methode vergeben."
    End If

End Sub

Private Sub cmdAddProperty_Click()

    If txtAddPropertyName.value <> "" And cmbAddPropertyType.value <> "" Then
    
        lstPreviewProperties.AddItem txtAddPropertyName.value & ";" & _
            cmbAddPropertyType.Column(1)
            
        txtAddPropertyName.value = ""
        cmbAddPropertyType.value = ""
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht befüllt."
        
    End If


End Sub

Private Sub cmdCreateClass_Click()
'TO-DO: Klasse erstellen
'



End Sub

Private Sub Form_Load()

    DisableAllPages
    pagClassData.SetFocus
    Load_Packages
    ApplyDefaultSettings
    
    lstPreviewProperties.ColumnCount = 2
    lstPreviewMethods.ColumnCount = 3
    lstPackages.ColumnCount = 1
    
    ClearListBoxEntries Me.Name, lstPreviewMethods
    ClearListBoxEntries Me.Name, lstPreviewProperties
    
    UpdateListBoxNavigationButtons Me.Name, lstPreviewProperties, cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
    UpdateListBoxNavigationButtons Me.Name, lstPreviewMethods, cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
    

End Sub
Private Sub ApplyDefaultSettings()

    txtAddPropertyName = ""
    txtAddMethodName = ""
    txtClassName = ""
    
    optMethodAddTypeSub = False
    optMethodAddTypeFunction = False
    optMethodAddPrivate = False
    optMethodAddPublic = False
End Sub
Private Sub DisableAllPages()

    pagDraft.Visible = False

End Sub
'Seite 1 - Klassendaten
Private Sub Load_Packages()

      Dim intRow As Long
    Dim intCol As Long
    Dim intRows As Long
    Dim intCols As Long
    Dim varRow() As Variant
    Dim varData() As Variant

    lstPackages.RowSource = ""
    
    varData = Get_Array_FromQuery("110_qryClassBuilder_Package_SORT")

    On Error GoTo ExitSub
    intRows = UBound(varData, 1)
    intCols = UBound(varData, 2)
    ' ? 2D-Array erkannt

    For intRow = 1 To intRows
        ReDim varRow(0 To intCols - 1)
        For intCol = 1 To intCols
            varRow(intCol - 1) = varData(intRow, intCol)
        Next intCol
        lstPackages.AddItem varRow(0)
        For intCol = 1 To intCols - 1
'            lstPackages.List(lstPackages.ListCount - 1, intCol) = varRow(intCol)
        Next intCol
    Next intRow
    Exit Sub

ExitSub:
    ' Falls 1D-Array, wird hier weitergemacht
    On Error Resume Next
'    lstPackages.Clear
    For intRow = LBound(varData) To UBound(varData)
        lstPackages.AddItem varData(intRow)
    Next intRow

End Sub


Private Sub lstPackages_AfterUpdate()


    Dim Selected As Variant
    Dim rcsPropertiesCurrentPackage As Recordset
    Dim rcsMethodsCurrentPackage As Recordset
    Dim intCounterArray As Integer
    Dim PreviewProperties() As Variant
    Dim lngPackageID As Long
    
    lstPreviewProperties.RowSource = ""
    
    Listbox_Clear lstPreviewMethods
    Listbox_Clear lstPreviewProperties
    
    Selected = Get_Listbox_Selected(lstPackages)

    If Not IsEmpty(Selected) Then
    
        lstPreviewProperties.ColumnCount = 2
        lstPreviewMethods.ColumnCount = 3
    
        
        For intCounterArray = LBound(Selected) To UBound(Selected)
        
            lngPackageID = dlookup("ID", "110_tblClassBuilder_Package", "Name ='" & Selected(intCounterArray) & "'")
        
            'Eigenschaften hinzufügen
            Set rcsPropertiesCurrentPackage = CurrentDb.OpenRecordset( _
            "SELECT * FROM 110_tblClassBuilder_Property_Draft " & _
            "WHERE Package_FK = " & lngPackageID)
                
            If rcsPropertiesCurrentPackage.RecordCount > 0 Then
            
                rcsPropertiesCurrentPackage.MoveFirst
                
                Do
        
                    lstPreviewProperties.AddItem rcsPropertiesCurrentPackage.Fields("Name").value & ";" & _
                        dlookup("Name", "110_tblClassBuilder_Property_Type", " ID = " & rcsPropertiesCurrentPackage.Fields("Type_FK").value)
        
                    rcsPropertiesCurrentPackage.MoveNext
            
                Loop While rcsPropertiesCurrentPackage.EOF = False
            
            End If
            
            
            'Methoden hinzufügen
            Set rcsMethodsCurrentPackage = CurrentDb.OpenRecordset( _
            "SELECT * FROM 110_tblClassBuilder_Method_Draft " & _
            "WHERE Package_FK = " & lngPackageID)
            
            If rcsMethodsCurrentPackage.RecordCount > 0 Then
            
                rcsMethodsCurrentPackage.MoveFirst
                
                Do
        
                    lstPreviewMethods.AddItem _
                        rcsMethodsCurrentPackage.Fields("Name").value & ";" & _
                        dlookup("Name", "110_tblClassBuilder_Method_Type", "ID = " & rcsMethodsCurrentPackage.Fields("Type_FK").value) & ";" & _
                        dlookup("Name", "110_tblClassBuilder_Visability", "ID = " & rcsMethodsCurrentPackage.Fields("Visability_FK").value)
        
                    rcsMethodsCurrentPackage.MoveNext
            
                Loop While rcsMethodsCurrentPackage.EOF = False
                
            End If
        
        Next intCounterArray

    End If

End Sub
Public Sub Listbox_Clear(objListBox As Access.Listbox)

    ' Leert eine Access-Listbox unabhängig vom aktuellen RowSourceType

    
    If objListBox.RowSourceType = "Table/Query" Then
        objListBox.RowSource = ""
    ElseIf objListBox.RowSourceType = "Value List" Then
        objListBox.RowSource = ""
    End If

End Sub


Private Sub cmdPreviewMethod_MoveDown_Click()
    ListBox_Item_Move lstPreviewMethods, Down
End Sub
Private Sub cmdPreviewMethod_MoveUp_Click()
    ListBox_Item_Move lstPreviewMethods, Up
End Sub
Private Sub cmdPreviewProperty_MoveDown_Click()
    ListBox_Item_Move lstPreviewProperties, Down
End Sub
Private Sub cmdPreviewProperty_MoveUp_Click()
    ListBox_Item_Move lstPreviewProperties, Up
End Sub
Public Sub ListBox_Item_Move(objListBox As Listbox, Direction As enuDirection)

    ' Verschiebt die markierte Zeile in einer mehrspaltigen Access-ListBox (Value List) um eine Position
    ' Unterstützt beliebig viele Spalten – funktioniert nur mit RowSourceType = "Value List"

    
    Dim intCols As Integer
    Dim intRows As Integer
    Dim intIndex As Integer
    Dim varData() As Variant
    Dim i As Long, j As Long
    Dim strRowSource As String
    Dim varTemp() As String

    If objListBox.RowSourceType <> "Value List" Then
        MsgBox "Diese Funktion unterstützt nur ListBoxen mit RowSourceType = 'Value List'", vbExclamation
        Exit Sub
    End If

    intCols = objListBox.ColumnCount
    intRows = objListBox.ListCount

    ' Auswahl finden
    intIndex = -1
    For i = 0 To intRows - 1
        If objListBox.Selected(i) Then
            intIndex = i
            Exit For
        End If
    Next i

    If intIndex = -1 Then Exit Sub ' nichts ausgewählt
    If Direction = enuDirection.Up And intIndex = 0 Then Exit Sub
    If Direction = enuDirection.Down And intIndex = intRows - 1 Then Exit Sub

    ' Daten in Array kopieren
    ReDim varData(0 To intRows - 1, 0 To intCols - 1)
    For i = 0 To intRows - 1
        For j = 0 To intCols - 1
            varData(i, j) = objListBox.Column(j, i)
        Next j
    Next i

    ' Zeilen tauschen
    Dim rowA As Long, rowB As Long
    If Direction = enuDirection.Up Then
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

    ' Neue RowSource erzeugen
    strRowSource = ""
    For i = 0 To intRows - 1
        ReDim varTemp(0 To intCols - 1)
        For j = 0 To intCols - 1
            varTemp(j) = Nz(varData(i, j), "")
        Next j
        strRowSource = strRowSource & Join(varTemp, ";") & ";"
    Next i

    ' Letztes Semikolon entfernen
    If Right(strRowSource, 1) = ";" Then
        strRowSource = Left(strRowSource, Len(strRowSource) - 1)
    End If

    objListBox.RowSource = strRowSource

    ' Neuen Eintrag wieder markieren
    objListBox.Selected(rowB) = True
    
    UpdateListBoxNavigationButtons Me.Name, lstPreviewMethods, cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
    UpdateListBoxNavigationButtons Me.Name, lstPreviewProperties, cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
    
End Sub

'#################################### Listbox - Auswahl #############################################
Private Sub lstPreviewMethods_Click()
    UpdateListBoxNavigationButtons Me.Name, lstPreviewMethods, cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreviewMethods_GotFocus()
    UpdateListBoxNavigationButtons Me.Name, lstPreviewMethods, cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreviewProperties_Click()
    UpdateListBoxNavigationButtons Me.Name, lstPreviewProperties, cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
Private Sub lstPreviewProperties_GotFocus()
    UpdateListBoxNavigationButtons Me.Name, lstPreviewProperties, cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub

Public Sub UpdateListBoxNavigationButtons( _
    strFormName As String, _
    objListBox As Object, _
    objButtonUp As Object, _
    objButtonDown As Object)

    ' Aktiviert oder deaktiviert die Buttons zum Verschieben eines Listbox-Eintrags je nach Auswahlposition
    ' Der Zugriff auf die Steuerelemente erfolgt direkt über Forms(strFormName).Controls("...").Enabled

    
    Dim intIndex As Long
    Dim intCount As Long

    If Not CurrentProject.AllForms(strFormName).IsLoaded Then Exit Sub

    intCount = Forms(strFormName).Controls(objListBox.Name).ListCount
    intIndex = Forms(strFormName).Controls(objListBox.Name).ListIndex + 1

    ' Wenn nichts ausgewählt oder Liste leer ? Buttons deaktivieren
    If intCount = 0 Or Forms(strFormName).Controls(objListBox.Name).ListIndex = -1 Then
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

Private Sub optMethodAddPrivate_AfterUpdate()
    SyncOptionFields Me.Name, optMethodAddPrivate, optMethodAddPublic
End Sub
Private Sub optMethodAddPublic_AfterUpdate()
    SyncOptionFields Me.Name, optMethodAddPublic, optMethodAddPrivate
End Sub
Private Sub optMethodAddTypeFunction_AfterUpdate()
    SyncOptionFields Me.Name, optMethodAddTypeFunction, optMethodAddTypeSub
End Sub
Private Sub optMethodAddTypeSub_AfterUpdate()
    SyncOptionFields Me.Name, optMethodAddTypeSub, optMethodAddTypeFunction
End Sub
Public Sub SyncOptionFields( _
    strFormName As String, _
    objChangedOption As Object, _
    objOtherOption As Object)

    ' Wenn das geänderte Optionsfeld aktiviert wurde,
    ' wird das andere automatisch deaktiviert

    
    Dim frm As Access.Form
    Set frm = Forms(strFormName)

    If frm.Controls(objChangedOption.Name).value = True Then
        frm.Controls(objOtherOption.Name).value = False
    Else
        frm.Controls(objOtherOption.Name).value = True
    End If

End Sub

