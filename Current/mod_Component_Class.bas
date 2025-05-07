Option Compare Database
Option Explicit

Dim Coding_VBA_Analyze As New clsCoding_VBA_Analyze

Public Sub InsertClassModules()
    ' Fügt alle Klassenmodule des aktuellen VBA-Projekts in die Tabelle tbl_Class ein
    ' Deaktiviert Klassen, die nicht mehr im Projekt vorhanden sind

    Dim dbDatabase As DAO.Database
    Dim rstClass As DAO.Recordset
    Dim objComponent As VBIDE.VBComponent
    Dim strClassName As String
    Dim lngExists As Long

    Set dbDatabase = CurrentDb
    Set rstClass = dbDatabase.OpenRecordset("tbl_Class", dbOpenDynaset)

    ' Alle vorhandenen Klassen zunächst inaktiv setzen

    rstClass.MoveFirst
    Do While Not rstClass.EOF

        If rstClass!Active = True Then
            rstClass.Edit
            rstClass!Active = False
            rstClass.Update
        End If

        rstClass.MoveNext
    Loop

    ' Aktuelle Klassenmodule hinzufügen bzw. reaktivieren und Properties aktualisieren

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents

        If objComponent.Type = vbext_ct_ClassModule Then

            strClassName = objComponent.Name

            ' Prüfen, ob die Klasse bereits in tbl_Class existiert
            lngExists = DCount("*", "tbl_Class", "Name = '" & strClassName & "'")

            If lngExists = 0 Then
                ' Neue Klasse einfügen und aktiv setzen
                rstClass.AddNew
                rstClass!Name = strClassName
                rstClass!Active = True
                rstClass.Update
            Else
                ' Bereits vorhandene Klasse wieder aktivieren
                rstClass.FindFirst "Name = '" & strClassName & "'"
                If Not rstClass.NoMatch Then
                    If rstClass!Active = False Then
                        rstClass.Edit
                        rstClass!Active = True
                        rstClass.Update
                    End If
                End If
            End If

            ' Properties für diese Klasse aktualisieren

            Update_Class_Properties_ByClass strClassName

        End If

    Next objComponent

    rstClass.Close
    Set rstClass = Nothing
    Set dbDatabase = Nothing

End Sub
Public Sub Update_Class_Properties_ByClass(ByVal strClassName As String)
    ' Aktualisiert die Einträge in tbl_Class_Property für ein Klassenmodul

    Dim dbDatabase As DAO.Database
    Dim rstProperties As DAO.Recordset
    Dim lngClassID As Long
    Dim varElements As Variant
    Dim lIndex As Long
    Dim strElementName As String
    Dim strReturnType As String
    Dim strCodeBlock As String
    Dim strElementType As String
    Dim lngExists As Long

    Set dbDatabase = CurrentDb

    ' Klassen-ID ermitteln
    lngClassID = Nz(DLookup("ID", "tbl_Class", "Name = '" & strClassName & "'"), 0)
    If lngClassID = 0 Then Exit Sub

    Set rstProperties = dbDatabase.OpenRecordset("tbl_Class_Property", dbOpenDynaset)

    ' Alle Properties dieser Klasse zunächst inaktiv setzen

    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE tbl_Class_Property SET Active = False"
    DoCmd.SetWarnings True

    ' Aktuelle Property-Codeblöcke des Moduls holen

    varElements = Coding_VBA_Analyze.Get_CodeElements_AsArray( _
        Coding_VBA_Analyze.Get_CodeModule_AsString(strClassName))

    For lIndex = LBound(varElements) To UBound(varElements)

        strElementType = Coding_VBA_Analyze.Get_CodeElement_Type(varElements(lIndex))
        If strElementType = "Property" Then

            strElementName = Coding_VBA_Analyze.GetCodeElement_Name_AsString(varElements(lIndex))
            strReturnType = Coding_VBA_Analyze.Get_CodeElement_ReturnType_AsString(varElements(lIndex))
            strCodeBlock = varElements(lIndex)

            ' Prüfen, ob Property bereits vorhanden ist
            lngExists = DCount("*", "tbl_Class_Property", _
                        "Class_FK = " & lngClassID & _
                        " AND Name = '" & strElementName & "'")

            If lngExists = 0 Then
                ' Neue Property einfügen und aktiv setzen
                rstProperties.AddNew
                rstProperties!Class_FK = lngClassID
                rstProperties!Name = strElementName
                rstProperties!DataType = strReturnType
                rstProperties!Code = strCodeBlock
                rstProperties!Active = True
                rstProperties.Update
            Else
                ' Bereits vorhandene Property aktualisieren und aktiv setzen
                rstProperties.FindFirst _
                    "Class_FK = " & lngClassID & _
                    " AND Name = '" & strElementName & "'"
                If Not rstProperties.NoMatch Then
                    rstProperties.Edit
                    rstProperties!DataType = strReturnType
                    rstProperties!Code = strCodeBlock
                    rstProperties!Active = True
                    rstProperties.Update
                End If
            End If

        End If
    Next lIndex

    rstProperties.Close
    Set rstProperties = Nothing
    Set dbDatabase = Nothing

End Sub


'
'Public Sub ExportAllClassProperties()
''To Refactor
'    ' Durchläuft alle Klassenmodule und gibt die Property-Namen jedes Moduls als formatierte String-Tabelle aus
'
'    Dim objComponent As VBIDE.VBComponent
'    Dim strClassName As String
'    Dim varProperties As Variant
'    Dim strOutput As String
'    Dim i As Long
'
'    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
'        If objComponent.Type = vbext_ct_ClassModule Then
'            strClassName = objComponent.Name
'            varProperties = Access_Class_Module_Get_PropertyNames(strClassName)
'
'            strOutput = strOutput & "Klasse: " & strClassName & vbCrLf
'
'            If Not IsNull(varProperties) Then
'                For i = LBound(varProperties) To UBound(varProperties)
'                    strOutput = strOutput & "  - " & varProperties(i) & vbCrLf
'                Next i
'            Else
'                strOutput = strOutput & "  (Keine Properties gefunden)" & vbCrLf
'            End If
'
'            strOutput = strOutput & vbCrLf
'        End If
'
'    Next objComponent
'
'    Debug.Print strOutput
'
'
'End Sub
Public Sub ExportAllClassMethods()
'To Refactor
    ' Durchläuft alle Klassenmodule und gibt alle Methodennamen (Sub/Function) jedes Moduls formatiert aus

    Dim objComponent As VBIDE.VBComponent
    Dim strClassName As String
    Dim varMethods As Variant
    Dim strOutput As String
    Dim i As Long

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        If objComponent.Type = vbext_ct_ClassModule Then
            strClassName = objComponent.Name
            varMethods = Access_Class_Module_Get_MethodNames(strClassName)

            strOutput = strOutput & "Klasse: " & strClassName & vbCrLf

            If Not IsNull(varMethods) Then
                For i = LBound(varMethods) To UBound(varMethods)
                    strOutput = strOutput & "  - " & varMethods(i) & vbCrLf
                Next i
            Else
                strOutput = strOutput & "  (Keine Methoden gefunden)" & vbCrLf
            End If

            strOutput = strOutput & vbCrLf
        End If
    Next objComponent

    Debug.Print strOutput

    
End Sub