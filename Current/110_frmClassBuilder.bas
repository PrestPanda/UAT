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

Dim Log As New clsLog

Dim arrSelectedPackages As String

Private Sub Form_Load()

    Log.WriteLine "Class Builder geöffnet."

    
    pagClassData.SetFocus
    Load_Packages
    ApplyDefaultSettings

    
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewProperties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewMethods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
    

End Sub
Private Sub Form_Close()

    Log.WriteLine "Class Builder geschlossen."

End Sub
Private Sub ApplyDefaultSettings()

    txtAddPropertyName = ""
    txtAddMethodName = ""
    txtClassName = ""
    
    optMethodAddTypeSub = False
    optMethodAddTypeFunction = False
    optMethodAddPrivate = False
    optMethodAddPublic = False
    
    lstPreviewProperties.ColumnCount = 2
    lstPreviewMethods.ColumnCount = 3
    lstPackages.ColumnCount = 1
    
    Access_ListBox_Clear Me.Name, "lstPreviewMethods"
    Access_ListBox_Clear Me.Name, "lstPreviewProperties"
    
End Sub
Private Sub cmdAddProperty_Click()

    If txtAddPropertyName.Value <> "" And cmbAddPropertyType.Value <> "" Then
    
        If IsNull(DLookup("ID", "110_tblClassBuilder_Property_Draft", "Name = '" & txtAddPropertyName.Value & "'")) = False Then
            MsgBox "Der Name der Property ist bereits an eine andere Property vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Property an und probieren Sie es erneut."
            Exit Sub
        End If
     
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreviewProperties", 0, txtAddPropertyName.Value) = False Then

            lstPreviewProperties.AddItem txtAddPropertyName.Value & ";" & _
                cmbAddPropertyType.Column(1)
                
            txtAddPropertyName.Value = ""
            cmbAddPropertyType.Value = ""
            
            txtAddPropertyName.SetFocus
            
        Else
        
            MsgBox "Eine Eigenschaft mit dem gleichen Namen wurde bereits hinzugefügt."
        
        End If
        
    Else
        
        MsgBox "Eines der Pflichtfelder wurde nicht befüllt."
        
    End If


End Sub
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
    
        If IsNull(DLookup("ID", "110_tblClassBuilder_Method_Draft", "Name = '" & txtAddMethodName & "'")) = False Then
            MsgBox "Der Name der Methode ist bereits an eine andere Methode vergeben worden, die Inhalt eines Pakets ist." & vbNewLine & _
                "Bitte passen Sie den Namen der Methode an und probieren Sie es erneut."
            Exit Sub
        End If
    
        If Access_ListBox_ContainsValue_InColumn(Me.Name, "lstPreviewMethods", 0, txtAddMethodName) = False Then
    
            If optMethodAddPrivate = True Then strVisability = "Private"
            If optMethodAddPublic = True Then strVisability = "Public"
            
            If optMethodAddTypeFunction = True Then strType = "Function"
            If optMethodAddTypeSub = True Then strType = "Sub"
    
            
             lstPreviewMethods.AddItem txtAddMethodName.Value & ";" & _
                strType & ";" & strVisability
                
            txtAddMethodName.Value = ""
            ApplyDefaultSettings
            
            txtAddMethodName.SetFocus
        
        Else
        
            MsgBox "Eine Methode mit diesem Namen wurde bereits hinzugefügt."
        
        End If
        
    Else
    
        MsgBox "Es wurde kein Name für die Methode vergeben."
    
    End If

End Sub
Private Sub cmdCreateClass_Click()

    Dim Class As New clsClass

    Dim Properties() As Variant
    Dim Methods() As Variant
    Dim Classes() As Variant
    
    Properties = Access_ListBox_Get_Array(Me.Name, "lstPreviewProperties")
    Methods = Access_ListBox_Get_Array(Me.Name, "lstPreviewMethods")
    Classes = Access_ListBox_Get_Array(Me.Name, "lstPackage_Class_Required")
    
    Class.Build Me.txtClassName.Value, Properties(), Methods(), Classes()
    


End Sub
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

    Dim Packages As Variant
    Dim lngCurrentPackage_ID As Long
    Dim strCurrentPackage_Name As String
    Dim blnCurrentPackage_Selected As Boolean

    Dim lngCounter_Packages As Long
    Dim rcsPackage_Class_Required As Recordset
    Dim rcsPackage_Properties As Recordset
    Dim rcsPackage_Methods As Recordset
    
    Packages = Access_ListBox_Get_Array(Me.Name, "lstPackages")


    If Not IsEmpty(Packages) Then
    
        For lngCounter_Packages = LBound(Packages) To UBound(Packages)
        
            strCurrentPackage_Name = Packages(lngCounter_Packages, 1)
            lngCurrentPackage_ID = DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & strCurrentPackage_Name & "'")
            blnCurrentPackage_Selected = Access_ListBox_IsValueSelected(Me.Name, "lstPackages", 0, strCurrentPackage_Name)
            
            'To-Do: Abhängigkeiten zu anderen Classen
            Set rcsPackage_Class_Required = CurrentDb.OpenRecordset( _
                "SELECT * FROM 110_tblClassBuilder_Package_Class_Required " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID)
                
            If rcsPackage_Class_Required.RecordCount > 0 Then
            
                rcsPackage_Class_Required.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPackage_Class_Required", rcsPackage_Class_Required.Fields("ClassName").Value) = False Then
                            'Eintrag hinzufügen
                            lstPackage_Class_Required.AddItem rcsPackage_Class_Required.Fields("ClassName").Value
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPackage_Class_Required", rcsPackage_Class_Required.Fields("ClassName").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPackage_Class_Required", rcsPackage_Class_Required.Fields("ClassName").Value
                        End If
                    
                    End If
                    
                    
                    rcsPackage_Class_Required.MoveNext
                    
                Loop While rcsPackage_Class_Required.EOF = False
            
            End If


            
            'Eigenschaften des Pakets
            Set rcsPackage_Properties = CurrentDb.OpenRecordset( _
                "SELECT * FROM 110_tblClassBuilder_Property_Draft " & _
                "WHERE Package_FK = " & lngCurrentPackage_ID)
                
            If rcsPackage_Properties.RecordCount > 0 Then
            
                rcsPackage_Properties.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewProperties", rcsPackage_Properties.Fields("Name").Value) = False Then
                            'Eintrag hinzufügen
                            lstPreviewProperties.AddItem rcsPackage_Properties.Fields("Name").Value & ";" & _
                                DLookup("name", "110_tblClassBuilder_Property_Type", "ID = " & rcsPackage_Properties.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewProperties", rcsPackage_Properties.Fields("Name").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPreviewProperties", rcsPackage_Properties.Fields("Name").Value
                        End If
                    
                    End If
                       
                    rcsPackage_Properties.MoveNext
                
                Loop While rcsPackage_Properties.EOF = False
                
            
            End If
            
            
            
            'Methoden des Pakets
            Set rcsPackage_Methods = CurrentDb.OpenRecordset("SELECT * FROM 110_tblClassBuilder_Method_Draft " & _
                "WHERE Package_FK = " & _
                DLookup("ID", "110_tblClassBuilder_Package", "Name = '" & strCurrentPackage_Name & "'"))
                
            If rcsPackage_Methods.RecordCount > 0 Then
            
                rcsPackage_Methods.MoveFirst
                
                Do
                
                    If blnCurrentPackage_Selected = True Then
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewMethods", rcsPackage_Methods.Fields("Name").Value) = False Then
                            lstPreviewMethods.AddItem rcsPackage_Methods.Fields("Name").Value & ";" & _
                                DLookup("name", "110_tblClassBuilder_Visability", "ID = " & rcsPackage_Methods.Fields("Visability_FK").Value) & ";" & _
                                DLookup("name", "110_tblClassBuilder_Method_Type", "ID = " & rcsPackage_Methods.Fields("Type_FK").Value)
                        End If
                        
                    Else
                    
                        If Access_ListBox_ContainsValue(Me.Name, "lstPreviewMethods", rcsPackage_Methods.Fields("Name").Value) = True Then
                            'Eintrag löschen
                            Access_ListBox_RemoveValue Me.Name, "lstPreviewMethods", rcsPackage_Methods.Fields("Name").Value
                        End If
                    
                    End If
                       
                    rcsPackage_Methods.MoveNext
                
                Loop While rcsPackage_Methods.EOF = False
                
            
            End If
        
        Next lngCounter_Packages
        
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


'#################################### Listbox - Auswahl #############################################
Private Sub lstPreviewMethods_Click()
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewMethods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreviewMethods_GotFocus()
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewMethods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
End Sub
Private Sub lstPreviewProperties_Click()
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewProperties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub
Private Sub lstPreviewProperties_GotFocus()
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewProperties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
End Sub

Public Sub UpdateListBoxNavigationButtons( _
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

    If frm.Controls(objChangedOption.Name).Value = True Then
        frm.Controls(objOtherOption.Name).Value = False
    Else
        frm.Controls(objOtherOption.Name).Value = True
    End If

End Sub

Public Sub ListBox_Item_Move(objListBox As ListBox, Direction As enuDirection)
'To-Do: Refactor
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
    
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewMethods", cmdPreviewMethod_MoveUp, cmdPreviewMethod_MoveDown
    UpdateListBoxNavigationButtons Me.Name, "lstPreviewProperties", cmdPreviewProperty_MoveUp, cmdPreviewProperty_MoveDown
    
End Sub
