Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Log As New clsLog
Dim objShell As Object

Public Sub CommitPushCurrentComponents(strCommitMessage As String)

    StagingBranch_AddAll
    Commit strCommitMessage
    Push

End Sub

Private Sub StagingBranch_AddAll()

    RunCommand "add ."
    Log.WriteLine "GIT - Staging Branch - All Added"

End Sub
Private Sub Commit(strCommitMessage As String)

     If strCommitMessage = "" Then
        RunCommand "commit --allow-empty-message -m """""
    Else
        RunCommand "commit -m """ & strCommitMessage & """"
    End If

    Log.WriteLine "GIT - Comitted"

End Sub
Private Sub Push()

    RunCommand "push origin master"
    
    Log.WriteLine "GIT - Pushed"

End Sub
Private Sub RunCommand(strCommand As String)

    objShell.Run "cmd.exe /c " & "cd /d """ & Config_Git_PathRepo & """ && git " & strCommand, 1, True

End Sub

Private Sub Class_Initialize()

    Set objShell = CreateObject("WScript.Shell")

End Sub

Private Sub Class_Terminate()

    Set objShell = Nothing

End Sub