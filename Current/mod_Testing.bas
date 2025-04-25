Option Compare Database
Option Explicit

Private Sub Testing_Created_Class_clsUser()

    Dim User As New clsUser
    
    With User
        .Name = "Jan"
        .IsActive = True
        .CreateUser_FK = 1
        .CreateTS = Now()
        .Save
        .Reset
    End With


End Sub