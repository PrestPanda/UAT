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

Private Sub cmdAddProperty_Click()



End Sub

Private Sub Form_Load()

    DisableAllPages
    pagClassData.SetFocus
    Load_Packages

End Sub
Private Sub DisableAllPages()

    pagDraft.Visible = False

End Sub
'Seite 1 - Klassendaten

Private Sub cmdConfirm_Click()

    Me.pagDraft.Visible = True
    regSteps = 1

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

'ListBox
'Herkunftstyp: Wertliste
'Mehrfachauswahl: Einzeln
Private Sub Update_Previews()


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

