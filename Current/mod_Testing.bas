Option Compare Database
Option Explicit

Private Sub Testing_Created_Class_clsUser()
'
    Dim User As New clsUser
    Dim User1 As New clsUser

    With User
        .Name = "Jan Christeleit"
        .Save
        .Reset
        .Name = "Marcel Kruse"
        .Save
        .Name = "Marcel Duse"
        .Save
        .Delete
        .LoadByID 1
    End With
    
    User1.ID = 1
    User1.Load
    User1.Delete


End Sub