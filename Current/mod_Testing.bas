Option Compare Database
Option Explicit

Private Sub Testing_Created_Class_clsUser()
'
    Dim User As New clsUser
    Dim User1 As New clsUser

    With User
        .Name = "Jan Christeleit"
        .IsActive = True
        .CreateUser_FK = 1
        .CreateTS = Now()
        .Save
        .Reset
        .Name = "Marcel Kruse"
        .IsActive = True
        .CreateUser_FK = 1
        .Save
    End With
    
    User1.ID = 1
    User1.Load
    User1.Delete


End Sub