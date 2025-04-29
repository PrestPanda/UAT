Option Compare Database
Option Explicit

Public Sub InsertClassModules()

    ' Fügt alle Klassenmodule des aktuellen VBA-Projekts in die Tabelle tblComponents ein

    Dim objComponent As VBIDE.VBComponent
    Dim dbDatabase As DAO.Database
    Dim strSQL As String
    Dim lngComponentTypeID As Long
    
    Set dbDatabase = CurrentDb
    lngComponentTypeID = 2 ' Annahme: 2 steht für Klassenmodule in tblComponentTypes

    For Each objComponent In Application.VBE.ActiveVBProject.VBComponents
        If objComponent.Type = vbext_ct_ClassModule Then
            
            strSQL = "INSERT INTO tbl_Class (Name) " & _
                     "VALUES ('" & objComponent.Name & "')"
                     
            DoCmd.SetWarnings False
                     
            DoCmd.RunSQL strSQL
            
            DoCmd.SetWarnings True
            
        End If
    Next objComponent

    
End Sub