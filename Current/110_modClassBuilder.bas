Option Compare Database
Option Explicit



Sub AddPropertyDraft_Standard()

    Dim rcsProperty_Standard As Recordset
    
    rcsProperty_Standard.OpenRecordset _
    ("SELECT * FROM 110_tblClassBuilder_Property_Draft " & _
     "WHERE Package_Name = 'Standard'")
     
     
    InsertPropertyTable rcsProperty_Standard

End Sub
Sub AddPropertyDraft_Tracking()

    Dim rcsProperty_Tracking As Recordset
    
    rcsProperty_Standard.OpenRecordset _
    ("SELECT * FROM 110_tblClassBuilder_Property_Draft " & _
     "WHERE Package_Name = 'Tracking'")
     
     
    InsertPropertyTable rcsProperty_Tracking


    

End Sub
Sub InsertPropertyTable(rcsProperty_Draft As Recordset)

    rcsProperty_Draft.MoveLast
     
     If rcsProperty_Draft.RecordCount > 0 Then
     
        rcsProperty_Draft.MoveFirst
        
        Do
        
            DoCmd.RunSQL _
                "INSERT INTO 110_ClassBuilder_Property (Name, Type_FK, Order) " & _
                "VALUES ('" & rcsProperty_Draft.Fields("Name").value & "'," & _
                rcsProperty_Draft.Fields("Type_FK").value & ",0)"
        
        Loop While rcsProperty_Draft.EOF = False
     
     End If

End Sub