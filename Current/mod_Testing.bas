Option Compare Database
Option Explicit

Dim Coding_VBA_Analyze As New clsCoding_VBA_Analyze


Public Sub Update_All()
    
    Update_Class

End Sub
Public Sub Update_Class()
    
    Update_Class_Properties

End Sub
Public Sub Update_Class_Properties()
    ' Aktualisiert die tbl_Class und tbl_Class_Property mit Klassenmodulen und deren Properties

    Dim dbDatabase As DAO.Database
    Dim rstClass As DAO.Recordset
    Dim rstProperties As DAO.Recordset
    Dim objComponent As VBIDE.VBComponent
    Dim strClassName As String
    Dim lngClassID As Long
    Dim varElements() As Variant
    Dim lIndex As Long
    Dim strElementName As String
    Dim strReturnType As String
    Dim strCodeBlock As String
    Dim strElementType As String

    Set dbDatabase = CurrentDb
    Set rstClass = dbDatabase.OpenRecordset("tbl_Class", dbOpenDynaset)
    Set rstProperties = dbDatabase.OpenRecordset("tbl_Class_Property", dbOpenDynaset)

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents

        If objComponent.Type = vbext_ct_ClassModule Then

            strClassName = objComponent.Name

            ' Prüfen, ob Klasse bereits vorhanden
            rstClass.FindFirst "Name = '" & strClassName & "'"
            If rstClass.NoMatch Then
                rstClass.AddNew
                rstClass!Name = strClassName
                rstClass.Update
                rstClass.FindFirst "Name = '" & strClassName & "'"
            End If

            lngClassID = rstClass!ID

            ' Alle Code-Elemente des Moduls holen
            varElements() = Coding_VBA_Analyze.Get_CodeElements_AsArray( _
                Coding_VBA_Analyze.Get_CodeModule_AsString(strClassName))

            For lIndex = LBound(varElements) To UBound(varElements)
                strElementType = Coding_VBA_Analyze.Get_CodeElement_Type(varElements(lIndex))
                If strElementType = "Property" Then
                    strElementName = Coding_VBA_Analyze.GetCodeElement_Name_AsString(varElements(lIndex))
                    strReturnType = Coding_VBA_Analyze.Get_CodeElement_ReturnType_AsString(varElements(lIndex))
                    strCodeBlock = varElements(lIndex)

                    ' Prüfen, ob Property bereits vorhanden
                    rstProperties.FindFirst "Class_FK = " & lngClassID & " AND Name = '" & strElementName & "'"
                    If rstProperties.NoMatch Then
                        rstProperties.AddNew
                        rstProperties!Class_FK = lngClassID
                        rstProperties!Name = strElementName
                        rstProperties!DataType = strReturnType
                        rstProperties!IsAddedInCode = 1
                        rstProperties!Code = strCodeBlock
                        rstProperties!Active = 1
                        rstProperties.Update
                    Else
                        rstProperties.Edit
                        rstProperties!DataType = strReturnType
                        rstProperties!IsAddedInCode = 1
                        rstProperties!Code = strCodeBlock
                        rstProperties!Active = 1
                        rstProperties.Update
                    End If
                End If
            Next lIndex

        End If

    Next objComponent

    rstClass.Close
    rstProperties.Close
    Set rstClass = Nothing
    Set rstProperties = Nothing
    Set dbDatabase = Nothing


End Sub

'Public Sub Update_Class_Properties()
'
'    ' Aktualisiert die Tabellen tbl_Class und tbl_Class_Property mit Klassenmodulen und deren Properties
'
'    Dim dbDatabase As DAO.Database
'    Dim rstClass As DAO.Recordset
'    Dim rstProperties As DAO.Recordset
'    Dim objComponent As VBIDE.VBComponent
'    Dim strClassName As String
'    Dim lngClassID As Long
'    Dim varProperties As Variant
'    Dim i As Long
'    Dim strPropertyName As String
'    Dim strPropertyType As String
'
'    Set dbDatabase = CurrentDb
'    Set rstClass = dbDatabase.OpenRecordset("tbl_Class", dbOpenDynaset)
'    Set rstProperties = dbDatabase.OpenRecordset("tbl_Class_Property", dbOpenDynaset)
'
'    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
'
'        If objComponent.Type = vbext_ct_ClassModule Then
'
'            strClassName = objComponent.Name
'
'            ' Prüfen, ob Klasse bereits vorhanden
'            rstClass.FindFirst "Name = '" & strClassName & "'"
'            If rstClass.NoMatch Then
'                rstClass.AddNew
'                rstClass!Name = strClassName
'                rstClass.Update
'                rstClass.FindFirst "Name = '" & strClassName & "'" ' wiederholen, um ID zu lesen
'            End If
'
'            lngClassID = rstClass!ID
'
'            varProperties = Access_Class_Module_Get_(strClassName)
'
'            If Not IsNull(varProperties) Then
'                For i = LBound(varProperties, 1) To UBound(varProperties, 1)
'
'                    strPropertyName = varProperties(i, 0)
'                    strPropertyType = varProperties(i, 1)
'
'                    ' Prüfen, ob die Property bereits vorhanden ist
'                    rstProperties.FindFirst "Class_FK = " & lngClassID & " AND Name = '" & strPropertyName & "'"
'
'                    If rstProperties.NoMatch Then
'                        rstProperties.AddNew
'                        rstProperties!Class_FK = lngClassID
'                        rstProperties!Name = strPropertyName
'                        rstProperties!DataType = strPropertyType
'                        rstProperties.Update
'                    End If
'
'                Next i
'            End If
'
'        End If
'
'    Next objComponent
'
'    rstClass.Close
'    rstProperties.Close
'    Set rstClass = Nothing
'    Set rstProperties = Nothing
'    Set dbDatabase = Nothing
'
'End Sub

'Public Sub Update_Class_Properties()
'
'    ' Aktualisiert die Tabellen tbl_Class und tbl_Class_Property mit Klassenmodulen und deren Properties
'
'    Dim dbDatabase As DAO.Database
'    Dim rstClass As DAO.Recordset
'    Dim rstProperties As DAO.Recordset
'    Dim objComponent As VBIDE.VBComponent
'    Dim strClassName As String
'    Dim lngClassID As Long
'    Dim varProperties As Variant
'    Dim i As Long
'    Dim strPropertyName As String
'    Dim strPropertyType As String
'
'    Set dbDatabase = CurrentDb
'    Set rstClass = dbDatabase.OpenRecordset("tbl_Class", dbOpenDynaset)
'    Set rstProperties = dbDatabase.OpenRecordset("tbl_Class_Property", dbOpenDynaset)
'
'    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
'
'        If objComponent.Type = vbext_ct_ClassModule Then
'
'            strClassName = objComponent.Name
'
'            ' Prüfen, ob Klasse bereits vorhanden
'            rstClass.FindFirst "Name = '" & strClassName & "'"
'            If rstClass.NoMatch Then
'                rstClass.AddNew
'                rstClass!Name = strClassName
'                rstClass.Update
'                rstClass.FindFirst "Name = '" & strClassName & "'" ' wiederholen, um ID zu lesen
'            End If
'
'            lngClassID = rstClass!ID
'
'            varProperties = Access_Class_Module_Get_PropertyNamesAndTypes(strClassName)
'
'            If Not IsNull(varProperties) Then
'                For i = LBound(varProperties) To UBound(varProperties)
'
'                    strPropertyName = varProperties(i)
'
'                    ' Prüfen, ob die Property bereits vorhanden ist
'                    rstProperties.FindFirst "Class_FK = " & lngClassID & " AND Name = '" & strPropertyName & "'"
'
'                    If rstProperties.NoMatch Then
'                        ' Datentyp ermitteln
'                        strPropertyType = Access_Class_Module_Get_PropertyType(strClassName, strPropertyName)
'
'                        rstProperties.AddNew
'                        rstProperties!Class_FK = lngClassID
'                        rstProperties!Name = strPropertyName
'                        rstProperties!DataType = strPropertyType
'                        rstProperties.Update
'                    End If
'
'                Next i
'            End If
'
'        End If
'
'    Next objComponent
'
'    rstClass.Close
'    rstProperties.Close
'    Set rstClass = Nothing
'    Set rstProperties = Nothing
'    Set dbDatabase = Nothing
'
'End Sub
'
'
'Public Sub Update_Class_Properties()
'
'    ' Aktualisiert die Tabellen tbl_Class und tbl_Class_Property mit Klassenmodulen und deren Properties
'
'    Dim dbDatabase As DAO.Database
'    Dim rstClass As DAO.Recordset
'    Dim rstProperties As DAO.Recordset
'    Dim objComponent As VBIDE.VBComponent
'    Dim strClassName As String
'    Dim lngClassID As Long
'    Dim varProperties As Variant
'    Dim i As Long
'
'    Set dbDatabase = CurrentDb
'    Set rstClass = dbDatabase.OpenRecordset("tbl_Class", dbOpenDynaset)
'    Set rstProperties = dbDatabase.OpenRecordset("tbl_Class_Property", dbOpenDynaset)
'
'    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
'
'        If objComponent.Type = vbext_ct_ClassModule Then
'
'            strClassName = objComponent.Name
'
'            ' Prüfen, ob Klasse bereits vorhanden
'            rstClass.FindFirst "Name = '" & strClassName & "'"
'            If rstClass.NoMatch Then
'                rstClass.AddNew
'                rstClass!Name = strClassName
'                rstClass.Update
'                rstClass.FindFirst "Name = '" & strClassName & "'" ' wiederholen, um ID zu lesen
'            End If
'
'            lngClassID = rstClass!ID
'
'            varProperties = Access_Class_Module_Get_PropertyNames(strClassName)
'
'            If Not IsNull(varProperties) Then
'                For i = LBound(varProperties) To UBound(varProperties)
'                    rstProperties.AddNew
'                    rstProperties!Class_FK = lngClassID
'                    rstProperties!Name = varProperties(i)
'                    rstProperties.Update
'                Next i
'            End If
'
'        End If
'
'    Next objComponent
'
'    rstClass.Close
'    rstProperties.Close
'    Set rstClass = Nothing
'    Set rstProperties = Nothing
'    Set dbDatabase = Nothing
'
'
'End Sub