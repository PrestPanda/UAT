Option Compare Database
Option Explicit


Public Sub ExportAllClassProperties()

    ' Durchläuft alle Klassenmodule und gibt die Property-Namen jedes Moduls als formatierte String-Tabelle aus

    Dim objComponent As VBIDE.VBComponent
    Dim strClassName As String
    Dim varProperties As Variant
    Dim strOutput As String
    Dim i As Long

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        If objComponent.Type = vbext_ct_ClassModule Then
            strClassName = objComponent.Name
            varProperties = Access_Class_Module_Get_PropertyNames(strClassName)
            
            strOutput = strOutput & "Klasse: " & strClassName & vbCrLf
            
            If Not IsNull(varProperties) Then
                For i = LBound(varProperties) To UBound(varProperties)
                    strOutput = strOutput & "  - " & varProperties(i) & vbCrLf
                Next i
            Else
                strOutput = strOutput & "  (Keine Properties gefunden)" & vbCrLf
            End If
            
            strOutput = strOutput & vbCrLf
        End If
        
    Next objComponent

    Debug.Print strOutput

    
End Sub
Public Sub ExportAllClassMethods()

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
Public Sub Update_All()
    
    Update_Class

End Sub
Public Sub Update_Class()
    
    Update_Class_Properties

End Sub
Public Sub Update_Class_Properties()

    ' Aktualisiert die Tabellen tbl_Class und tbl_Class_Property mit Klassenmodulen und deren Properties

    Dim dbDatabase As DAO.Database
    Dim rstClass As DAO.Recordset
    Dim rstProperties As DAO.Recordset
    Dim objComponent As VBIDE.VBComponent
    Dim strClassName As String
    Dim lngClassID As Long
    Dim varProperties As Variant
    Dim i As Long

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
                rstClass.FindFirst "Name = '" & strClassName & "'" ' wiederholen, um ID zu lesen
            End If
            
            lngClassID = rstClass!ID

            varProperties = Access_Class_Module_Get_PropertyNames(strClassName)

            If Not IsNull(varProperties) Then
                For i = LBound(varProperties) To UBound(varProperties)
                    rstProperties.AddNew
                    rstProperties!Class_FK = lngClassID
                    rstProperties!Name = varProperties(i)
                    rstProperties.Update
                Next i
            End If

        End If
    Next objComponent

    rstClass.Close
    rstProperties.Close
    Set rstClass = Nothing
    Set rstProperties = Nothing
    Set dbDatabase = Nothing

    
End Sub